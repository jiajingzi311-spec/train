# KV Cache Connector 策略分析

本文从策略视角总结 vLLM-Ascend 中 KV Cache Connector 的几类实现。为了便于选型，把当前实现分成两条主线：

- **Mooncake 主线**：以 Mooncake Transfer Engine 和 Mooncake Store 为核心，覆盖 P/D 分离 KV 传输、layerwise 传输、hybrid attention 传输，以及基于 Mooncake Store 的 KV Pool。
- **其他 connector 主线**：包括 UCM、LMCache-Ascend、AscendStore 的 Memcache/Yuanrong 后端，以及 CPU offload / recompute offload。它们更偏缓存层、外部存储层、生态集成或本机容量扩展。

这两条线不是互斥关系。vLLM-Ascend 当前的核心组合方式是 `MultiConnector`：让 Mooncake 负责 P/D 之间的热路径传输，让其他 connector 或 AscendStore 后端负责 prefix cache、持久化、外部存储或本机 offload。

## 总览

| 主线 | Connector / 后端 | 核心功能 | 典型场景 | Ascend 适配状态 |
| :--- | :--- | :--- | :--- | :--- |
| Mooncake | `MooncakeConnectorV1` | P/D 分离下 Decode 从 Prefill 拉取 KV | 标准 disaggregated prefill，单机或多机 P/D | 已注册到 vLLM-Ascend，配置 `kv_buffer_device: "npu"`，依赖 Mooncake Ascend Direct/Ascend transport |
| Mooncake | `MooncakeLayerwiseConnector` | Prefill 按层把 KV 推给 Decode，传输和计算重叠 | 追求更低 TTFT 的 P/D 分离 | 已注册，需要 layerwise proxy；和普通 Mooncake P2P 的调度链路不同 |
| Mooncake | `MooncakeHybridConnector` | 面向 hybrid attention、多 KV cache group、压缩 KV block 的 P/D KV 传输 | DeepSeek V4 Flash/Pro、Qwen hybrid attention 类模型 | 已注册，要求所有组合 connector 支持 HMA；当前有 TP/PP/CP 约束 |
| Mooncake | `AscendStoreConnector` + `backend: "mooncake"` | 分布式 KV Pool / prefix cache pool | PD-Mixed、P/D + KV Pool、跨实例前缀复用 | 已适配，Mooncake 后端要求 `protocol: "ascend"`；支持 FabricMem、SSD offload 等能力 |
| 其他 | `UCMConnector` | UCM 分层缓存、持久化 prefix cache、稀疏注意力、P/D 存储解耦 | 中心化 P/D、持久化 KV、长上下文、UCM + Mooncake P2P | vLLM-Ascend 提供封装，实际能力由外部 UCM 包提供 |
| 其他 | `LMCacheAscendConnector` | 使用 LMCache-Ascend 的动态 KV connector | 已采用 LMCache 生态、跨引擎 KV 复用、统一观测与存储后端 | vLLM-Ascend 侧是薄封装，依赖 `lmcache_ascend` |
| 其他 | `AscendStoreConnector` + `backend: "memcache"` | 以 Memcache/MemFabric 做 KV Pool | 简化部署的 KV Pool、layerwise KV Pool 试验 | 已适配；当前 layerwise KV Pool 主要落在 memcache 后端 |
| 其他 | `AscendStoreConnector` + `backend: "yuanrong"` | 以 Yuanrong Datasystem 做 KV Pool 后端 | 企业级数据系统、共享内存/远端 H2D 场景 | 已接入，需要 Yuanrong Datasystem、etcd、环境变量和可选 Remote H2D 环境 |
| 其他 | `OffloadingConnector` + `NPUOffloadingSpec` | NPU KV block 与 CPU block pool 之间异步换入换出 | NPU 显存紧张、长上下文、提高并发 | vLLM-Ascend 提供 NPU offloading spec，使用 NPU stream 做异步拷贝 |
| 其他 | `SimpleCPUOffloadConnector` | 覆盖上游 Simple CPU offload 的 NPU worker | 简单本机 CPU offload | 已覆盖注册，使用 `aclrtMemcpyBatchAsync` 和 `torch.npu` stream/event |
| 其他 | `RecomputeCPUOffloadConnector` | preempt/recompute 场景下保留 CPU KV | Decode 侧 recompute scheduler、抢占恢复 | 已注册；在 `AscendMultiConnector` 中对 preempted request 有优先级 |

## Mooncake 主线

### 1. 策略定位

Mooncake 的定位不是单一 connector，而是一套 KVCache-centric 的推理基础设施。它主要包含两层：

- **Transfer Engine**：面向大块或非连续内存的零拷贝/异步数据搬运抽象，核心概念是 `Segment` 和 `BatchTransfer`。在 GPU/NPU 集群上，它负责把 KV cache 从一个进程、设备或节点搬到另一个进程、设备或节点。
- **Mooncake Store**：面向 LLM KV cache 的分布式 KV 存储。它用 Master 维护元数据和空间分配，Client 同时可以作为请求端和存储段提供者；数据流不经过 Master，而是在 Client 之间直接传输。

因此，在 vLLM-Ascend 中 Mooncake 主线分成两个不同职责：

- `MooncakeConnectorV1`、`MooncakeLayerwiseConnector`、`MooncakeHybridConnector`：负责 **P/D 间的在线 KV 传输**。
- `AscendStoreConnector` + `backend: "mooncake"`：负责 **KV Pool / prefix cache pool**，把可复用 KV 放到 Mooncake Store 的全局资源池中。

### 2. `MooncakeConnectorV1`: 标准 P2P Pull

这是最基础的 P/D 分离策略：

1. Prefill 节点计算 prompt 的 KV。
2. Decode 节点在调度到该请求后，从 Prefill 节点拉取 KV。
3. Decode 加载 KV 后继续生成 token。

适合场景：

- P/D 资源拆分明确，Prefill 和 Decode 节点同构或网络链路稳定。
- 希望把长 prompt 的 prefill 压力从 decode 节点分离出去。
- 需要和 KV Pool、UCM prefix cache、CPU offload 等能力组合。

Ascend 使用要点：

- `kv_connector` 设置为 `MooncakeConnectorV1`。
- Prefill 使用 `kv_role: "kv_producer"`，Decode 使用 `kv_role: "kv_consumer"`。
- NPU 场景一般设置 `kv_buffer_device: "npu"`。
- Mooncake 配置文件里 `protocol` 需要使用 `ascend`。
- A3 上推荐优先使用 Ascend Direct / FabricMem 路径；老的 Ascend Transport 在 Mooncake 文档中已标注为后续弃用方向，Ascend 平台优先看 Ascend Direct Transport。

### 3. `MooncakeLayerwiseConnector`: 逐层 Push

`MooncakeLayerwiseConnector` 把一次完整 KV 传输拆成 layer 粒度。请求先进入 Decode，Decode 通过 layerwise proxy 的 metaserver 反向触发 Prefill；Prefill 一边计算，一边把每层 KV 推给 Decode。Decode 在每层需要 KV 时等待对应层的数据。

它的优势是 **把 Prefill 计算和 KV 传输重叠**，对长 prompt 的 TTFT 更友好。它适合：

- prompt 很长，完整 KV 拉取会成为明显等待点。
- P/D 间网络足够稳定，且希望把传输提前到 Prefill 前向过程中。
- 可以接受额外的 layerwise proxy 部署和排障成本。

限制和边界：

- 需要 `examples/disaggregated_prefill_v1/load_balance_proxy_layerwise_server_example.py` 这类 layerwise proxy。
- 它和 `AscendStoreConnector` 的 `use_layerwise: true` 不是同一件事。前者是 P2P 传输策略，后者是 KV Pool 的逐层读写优化。
- 在 `AscendMultiConnector` 中，`MooncakeLayerwiseConnector` 的调度状态会被特殊转发，因此和其他 connector 组合时要确认请求生命周期是否符合预期。

### 4. `MooncakeHybridConnector`: Hybrid Attention / HMA

`MooncakeHybridConnector` 面向 hybrid attention 模型。与普通 P2P 相比，它需要处理多 KV cache group、不同 block size、sliding window 或压缩 KV block 等问题。

适合场景：

- DeepSeek V4 Flash/Pro、Qwen hybrid attention 这类非均一 KV cache 布局。
- 模型启用 HMA，即 hybrid KV cache manager。
- P/D 分离仍然需要高性能 KV 传输，但普通 connector 无法正确表达多组 KV。

当前约束：

- 组合使用时，`AscendMultiConnector` 要求所有子 connector 都支持 HMA，否则不能处理多 KV cache group。
- 当前实现中对 PCP/DCP、prefill/decode TP、decode PP 等有硬约束；部署前要对照模型教程和源码确认。
- `AscendStoreConnector` 的 layerwise 模式当前不支持 hybrid KV cache groups，因此不要把它当成 hybrid attention 的通用 layerwise KV Pool 方案。

### 5. `AscendStoreConnector` + Mooncake Store

当 `AscendStoreConnector` 的 `backend` 设置为 `mooncake` 时，它使用 Mooncake Store 作为 KV Pool 后端。此时它不是直接的 P/D 传输通道，而是 prefix cache pool：

- Prefill 或 mixed 实例把已计算的 prefix KV 写入 Mooncake Store。
- 后续请求命中同样 prefix 时，从 pool 中加载 KV，减少重复 prefill。
- P/D 分离下常与 `MooncakeConnectorV1` 同时配置：P2P connector 负责本次请求的 P/D KV transfer，Store connector 负责跨请求复用。

Ascend 适配要点：

- `MOONCAKE_CONFIG_PATH` 指向 `mooncake.json`。
- `mooncake.json` 中 `protocol` 必须是 `ascend`。
- 支持通过 `MOONCAKE_MASTER`、`MOONCAKE_GLOBAL_SEGMENT_SIZE` 覆盖 master 地址和全局 segment size。
- A3 + FabricMem 场景可设置 `ASCEND_ENABLE_USE_FABRIC_MEM=1`；源码里该模式会让 Mooncake Store setup 使用 `local_buffer_size=0`。
- SSD offload 需要 Mooncake 版本和 `mooncake.json` 中 `enable_ssd_offload` / `ssd_offload_path` 配套。

### 6. Mooncake 主线的发展方向

Mooncake 线后续的重点大概率会落在几个方向：

- **更细粒度的重叠**：从整请求 KV 传输走向 layerwise、chunk-wise、甚至和流水线/attention 后端更紧密的重叠。
- **Hybrid KV 普适化**：让 hybrid attention、多 KV group、压缩 KV block、MLA、sparse attention 的元数据统一表达。
- **传输层升级**：Ascend 平台上从旧 Ascend Transport 迁移到 Ascend Direct Transport，增强 HCCS/RDMA、FabricMem、异步传输和超时控制。
- **存储层分层**：HBM/DRAM/FabricMem/SSD 的多层 KV Pool，并通过 Store 的 eviction、lease、replica、preferred segment 做稳定性治理。
- **组合编排**：通过 `MultiConnector` 和 HMA 支持，把 P2P、KV Pool、offload、recompute 组合成一套可恢复、可降级的策略。

## 其他 Connector 主线

### 1. `UCMConnector`: 统一缓存管理和持久化 KV

UCM 的策略重点是 **存储计算解耦**。它把 KV cache 从推理实例的临时状态提升为可持久化、可检索、可跨实例复用的数据层。典型层次是：

```text
HBM -> DRAM local cache -> Storage Backend
```

UCM 适合：

- 中心化 P/D：Prefill 把 KV 写到共享存储，Decode 从共享存储读取，调度更简单。
- 持久化 prefix cache：缓存可以跨服务重启、实例迁移或调度变化继续使用。
- 长上下文和稀疏注意力：UCM 官方文档还覆盖 GSA、CacheBlend 等训练无关稀疏/融合能力。
- 大规模专家并行 P/D：通过共享存储避免把 KV 状态绑定到单个推理实例。

Ascend 适配状态：

- vLLM-Ascend 中 `UCMConnector` 是对外部 `ucm.integration.vllm.ucm_connector.UCMConnector` 的封装。
- 官方 UCM 文档提供 vLLM-Ascend quickstart，安装时可设置 `PLATFORM=ascend`，A3 可使用 `PLATFORM=ascend-a3`。
- vLLM-Ascend 的 UCM 部署文档覆盖中心化 P/D 和 `MooncakeConnectorV1 + UCMConnector` 的分布式 P2P 组合。

与 Mooncake 的关系：

- 中心化 P/D 中可以只用 `UCMConnector`，不需要 Mooncake P2P。
- 分布式 P2P 中常见组合是 Prefill 侧 `MultiConnector` 同时配置 `MooncakeConnectorV1` 和 `UCMConnector`，Decode 侧只配置 `MooncakeConnectorV1`。此时 Mooncake 负责 P/D 热路径传输，UCM 负责 Prefill 侧 prefix cache。

后续方向：

- 更强的持久化缓存治理，包括生命周期、容量、命中率、缓存污染控制。
- 与稀疏注意力、CacheBlend、ReRoPE 等长上下文算法结合。
- 统一观测与调度反馈，让 cache hit、外部存储延迟、load failure 更容易被系统调度使用。

### 2. `LMCacheAscendConnector`: LMCache 生态接入

LMCache 的策略重点是 **跨引擎、跨实例、跨存储后端的 KV cache 管理层**。它强调 vendor-neutral 和 engine-independent，把 KV cache 持久化、复用、观测、压缩、传输等能力做成独立层。

适合：

- 已经使用 LMCache 或希望和 vLLM/SGLang 等多引擎生态统一。
- 需要丰富的存储后端，如 CPU RAM、本地 SSD、Redis/Valkey、S3、Mooncake Store、NIXL、GDS 等。
- 需要生产化观测、控制 API、KV 生命周期管理。
- 希望把未来的 KV 压缩、CacheBlend、非 prefix KV reuse 等能力纳入统一体系。

Ascend 适配状态：

- vLLM-Ascend 中 `LMCacheAscendConnector` 只是薄封装：导入 `lmcache_ascend` 后复用 vLLM 的 `LMCacheConnectorV1`。
- 需要单独安装 `LMCache-Ascend`，官方部署方式包括 Docker 和源码安装。
- 使用形态是 `kv_connector: "LMCacheAscendConnector"`，`kv_role: "kv_both"`。

与 Mooncake 的关系：

- LMCache 自身支持把 Mooncake Store 作为后端之一，这属于 **LMCache 体系内使用 Mooncake 存储后端**。
- vLLM-Ascend 直接同时配置 `LMCacheAscendConnector` 和 `AscendStoreConnector backend=mooncake` 时，要特别小心缓存所有权重复。一般应明确一个组件负责 prefix cache 命中与写入，另一个组件只负责传输或不用。
- 如果目标是纯 vLLM-Ascend P/D 热路径，Mooncake connector 更直接；如果目标是多引擎生态和统一 KV 管理，LMCache 线更合适。

后续方向：

- 与更多存储/传输后端融合，包括 Mooncake、NIXL、SSD/GDS、对象存储。
- KV cache 观测、配额、跨实例调度和运维控制能力增强。
- Ascend 插件从社区维护走向更稳定的版本矩阵和性能基准。

### 3. `AscendStoreConnector` 的 Memcache / Yuanrong 后端

`AscendStoreConnector` 本身是一层 vLLM-Ascend KV Pool 抽象，后端可插拔。当前源码里后端映射包括：

- `mooncake` -> `MooncakeBackend`
- `memcache` -> `MemcacheBackend`
- `yuanrong` -> `YuanrongBackend`

#### Memcache 后端

Memcache 后端适合希望使用 MemFabric/Memcache 体系做 KV Pool 的部署。它在文档中覆盖 P/D 分离和 PD-Mixed 两类场景，也支持 layerwise KV Pool。

使用场景：

- 希望有相对轻量的 KV Pool 后端。
- 需要 `AscendStoreConnector` 的 `use_layerwise: true`，并能接受当前主要落在 `backend: "memcache"` 的限制。
- 已经部署 MemFabric/Memcache 依赖。

边界：

- 需要独立的 meta/local config 和服务。
- layerwise KV Pool 和 `MooncakeLayerwiseConnector` 是两套机制，不能简单等价替换。
- 当前 layerwise KV Pool 不支持 hybrid KV cache groups。

#### Yuanrong 后端

Yuanrong 后端适合企业级数据系统或共享内存/远端 H2D 方向的 KV Pool。

使用场景：

- 已经有 Yuanrong Datasystem 或希望复用其服务发现、共享内存、worker-worker batch get 能力。
- 需要把 KV Pool 接到数据系统，而不是 Mooncake/Memcache。
- 有较强环境治理能力，可以配置 etcd、Datasystem worker、HugeTLB、NPU/RoCE 检查等。

边界：

- 部署复杂度高于 Mooncake/Memcache 简单路径。
- `DS_ENABLE_REMOTE_H2D=1` 需要满足 Remote H2D 的 NPU、驱动、HugeTLB 和网络要求。
- 更适合平台化或企业内部基础设施，不适合作为最小验证路径。

后续方向：

- `AscendStoreConnector` 的后端抽象会更像统一 KV Pool 接口，后端差异主要体现在存储分配、传输路径、故障恢复和可观测性。
- layerwise KV Pool 从 memcache 扩展到更多后端，是一个自然演进方向。
- Yuanrong 这类数据系统后端会更关注远端 H2D、共享内存、容量治理和多租户隔离。

### 4. CPU Offload / Recompute Offload

CPU offload 的策略与 Mooncake、UCM、LMCache 不同。它不是跨节点 KV 共享，也不是全局 prefix cache pool，而是把本机 NPU KV cache 的冷块换出到 CPU 内存。

当前相关实现包括：

- `OffloadingConnector` + `NPUOffloadingSpec`：基于 vLLM offloading 框架，vLLM-Ascend 提供 NPU 专用 spec。
- `SimpleCPUOffloadConnector`：vLLM-Ascend 覆盖上游注册，把 worker 替换成 NPU worker。
- `RecomputeCPUOffloadConnector`：面向 preempt/recompute 的请求状态保留。

适合：

- 单机或单实例 NPU KV 空间不足。
- 长上下文或高并发导致 NPU KV block 紧张。
- 希望 miss 时从 CPU 恢复，减少完全重算。
- Decode 侧抢占后，希望后续 recompute 能复用一部分 CPU KV。

Ascend 适配状态：

- `NPUOffloadingSpec` 使用专用 NPU stream 做 D2H/H2D 异步传输。
- `SimpleCPUOffloadConnector` 的 NPU worker 使用 `aclrtMemcpyBatchAsync` 和 `torch.npu` stream/event。
- `RecomputeCPUOffloadConnector` 在 `AscendMultiConnector` 中有优先级：如果某个请求已被 recompute offload 记录，优先从该 connector 查询可恢复 token。

后续方向：

- 与 upstream vLLM CPU offload 框架收敛，减少 Ascend 侧重复维护。
- 和 `kv_load_failure_policy: "recompute"`、preempt scheduler、KV events 结合，形成更完整的失败恢复路径。
- 与外部 KV Pool 分工清晰化：CPU offload 负责本机短期容量，外部 KV Pool 负责跨实例或持久化复用。

## Mooncake 与其他 Connector 的组合关系

### 组合原则

`MultiConnector` 是组合多个 connector 的入口，但它不是让多个 connector 同时对同一个请求产出传输结果。vLLM-Ascend 的 `AscendMultiConnector` 有几个关键规则：

- HMA 场景下，所有子 connector 都必须支持 HMA。
- `request_finished_all_groups` 中最多只能有一个 connector 产出 KV transfer params，否则会报错。
- 可以有多个异步保存任务，`AscendMultiConnector` 会记录额外 async save 数量。
- `RecomputeCPUOffloadConnector` 对已有 preempted request 有优先查询权。
- `MooncakeLayerwiseConnector` 在 `update_state_after_alloc` 中会被特殊转发状态。

因此，组合时最好按职责拆分：

- 一个 connector 负责 P/D 热路径传输。
- 一个 connector 负责 prefix cache / KV Pool / 持久化。
- 一个 connector 负责本机 offload 或 recompute 恢复。

### 推荐组合矩阵

| 组合 | 是否推荐 | 适用场景 | 注意点 |
| :--- | :--- | :--- | :--- |
| `MooncakeConnectorV1` + `AscendStoreConnector backend=mooncake` | 推荐 | P/D 分离 + Mooncake KV Pool | 常规官方路径；P2P 负责本次请求传输，Store 负责 prefix cache pool |
| `MooncakeConnectorV1` + `UCMConnector` | 推荐 | 分布式 P2P + UCM prefix cache | Prefill 侧用 `MultiConnector`，Decode 侧通常只配 Mooncake；UCM 负责 prefix cache，不接管 P/D 热路径 |
| `MooncakeConnectorV1` + `AscendStoreConnector backend=memcache` | 可用 | P/D 分离 + Memcache KV Pool | Mooncake 仍负责 P/D；Memcache 负责 pool；需要 Memcache 服务和 config |
| `MooncakeLayerwiseConnector` + KV Pool | 谨慎 | 低 TTFT P/D，同时希望 prefix cache pool | 区分 Mooncake layerwise P2P 和 Store layerwise KV Pool；检查 proxy、状态转发和后端限制 |
| `MooncakeHybridConnector` + Store/UCM/Offload | 谨慎 | hybrid attention 模型需要 P/D 和缓存组合 | 所有子 connector 需支持 HMA；避免使用不支持 hybrid KV groups 的 layerwise pool |
| `MooncakeConnectorV1` + CPU offload | 可用但需压测 | P/D 分离，同时本机 NPU KV 紧张 | 关注 NPU/CPU 拷贝与 P2P 传输是否争抢链路和调度窗口 |
| `MooncakeConnectorV1` + `RecomputeCPUOffloadConnector` | 可用但需验证 | P/D + preempt/recompute 恢复 | Recompute offload 对 preempted request 有优先级，需验证请求回滚和 load failure 策略 |
| `LMCacheAscendConnector` + Mooncake Store 后端 | 取决于架构 | LMCache 体系内使用 Mooncake 作为后端 | 推荐在 LMCache 内部选择 Mooncake 后端，而不是同时让 vLLM-Ascend 的 Store connector 和 LMCache 管同一份 prefix cache |
| `UCMConnector` 单独使用 | 推荐 | 中心化 P/D、持久化 KV、存储计算解耦 | 不追求最低 P2P 延迟时更简单；需要共享存储和 UCM 配置 |
| `AscendStoreConnector backend=yuanrong` 单独使用 | 平台化场景推荐 | 企业数据系统 KV Pool | 部署复杂，依赖 Datasystem、etcd、HugeTLB/NPU 网络环境 |

## 选型建议

| 目标 | 优先选择 |
| :--- | :--- |
| 标准 P/D 分离，追求热路径低延迟 | `MooncakeConnectorV1` |
| P/D 分离且长 prompt TTFT 压力大 | `MooncakeLayerwiseConnector` |
| Hybrid attention 模型做 P/D | `MooncakeHybridConnector` |
| P/D + prefix cache pool，且希望同一套 Mooncake 基础设施 | `MultiConnector` + `MooncakeConnectorV1` + `AscendStoreConnector backend=mooncake` |
| 中心化 P/D、持久化缓存、存储计算解耦 | `UCMConnector` |
| 已经采用 LMCache 生态，关注跨引擎和生产观测 | `LMCacheAscendConnector` |
| 轻量 KV Pool 或 layerwise KV Pool 验证 | `AscendStoreConnector backend=memcache` |
| 企业级数据系统和共享内存/远端 H2D | `AscendStoreConnector backend=yuanrong` |
| 单机 NPU KV 容量不足 | `OffloadingConnector` + `NPUOffloadingSpec` 或 `SimpleCPUOffloadConnector` |
| 抢占/recompute 场景下保留 KV | `RecomputeCPUOffloadConnector` |

## 后续演进判断

从当前代码和上游项目方向看，KV Cache Connector 的发展会沿着下面几条线推进：

1. **传输与存储分层会更清晰**：P/D 热路径传输、prefix cache pool、持久化存储、本机 CPU offload 会逐步形成明确职责边界。
2. **Connector 组合会成为常态**：单一 connector 覆盖所有场景的可能性不高，`MultiConnector` 的 HMA、失败恢复、async save 管理会越来越重要。
3. **Layerwise / chunk-wise 会继续下沉**：为了降低 TTFT，KV 传输和 attention 计算会更细粒度重叠。
4. **Hybrid attention 会推动元数据标准化**：多 KV group、不同 block size、压缩 KV、MLA、sparse attention 都要求 connector 共享更强的布局描述能力。
5. **Ascend 侧重点在 Direct/FabricMem/Remote H2D**：Mooncake Ascend Direct、FabricMem、Yuanrong Remote H2D、NPU stream offload 代表了不同层面的数据直达方向。
6. **持久化和可观测性会变成生产刚需**：UCM/LMCache 更强调缓存生命周期、命中率、容量、租户、指标和控制 API，这些能力会反向影响 vLLM-Ascend 的 connector 设计。

## 参考资料

- vLLM-Ascend KV Pool / Ascend Store: <https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/kv_pool.html>
- vLLM-Ascend UCM 部署: <https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/ucm_deployment.html>
- vLLM-Ascend LMCache-Ascend 部署: <https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/lmcache_ascend_deployment.html>
- vLLM-Ascend KV Cache CPU Offload: <https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/kv_cache_cpu_offload.html>
- Mooncake GitHub: <https://github.com/kvcache-ai/Mooncake>
- Mooncake Transfer Engine: <https://kvcache-ai.github.io/Mooncake/design/transfer-engine/index.html>
- Mooncake Store: <https://kvcache-ai.github.io/Mooncake/design/mooncake-store.html>
- Mooncake Ascend Direct Transport: <https://kvcache-ai.github.io/Mooncake/design/transfer-engine/ascend_direct_transport.html>
- Mooncake x LMCache Integration: <https://kvcache-ai.github.io/Mooncake/getting_started/examples/lmcache-integration.html>
- UCM 文档: <https://ucm.readthedocs.io/en/latest/>
- UCM vLLM-Ascend Quickstart: <https://ucm.readthedocs.io/en/latest/getting-started/quickstart_vllm_ascend.html>
- LMCache 文档: <https://docs.lmcache.ai/>
- LMCache GitHub: <https://github.com/LMCache/LMCache>
- LMCache-Ascend GitHub: <https://github.com/LMCache/LMCache-Ascend>
