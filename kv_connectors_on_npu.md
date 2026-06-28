# vLLM Ascend KV Connector 技术说明

本文整理 vLLM-Ascend 侧已经注册和适配的 KV Connector 用法，说明它们的功能边界、典型分类，以及在 Ascend NPU 上的部署配置方式。

## 范围

这里的 `kv-connect` 对应 vLLM V1 的 KV Connector 机制。vLLM-Ascend 在 `vllm_ascend/distributed/kv_transfer/__init__.py` 中注册或覆盖 connector，并通过 `--kv-transfer-config` 或 Python `KVTransferConfig` 启用。

需要注意两点：

- vLLM CLI 中仍沿用 `--gpu-memory-utilization` 等参数名，在 Ascend 场景下实际控制的是 NPU 侧缓存预算。
- `kv_role` 是最核心的角色开关：`kv_producer` 通常是 Prefill 节点，`kv_consumer` 通常是 Decode 节点，`kv_both` 用于同实例读写、KV Pool、UCM、LMCache、CPU offload 等场景。

## 实现分类总览

| 分类 | `kv_connector` 名称 | 主要实现 | 典型角色 | 核心能力 | 适用场景 |
| --- | --- | --- | --- | --- | --- |
| 组合编排 | `MultiConnector` | `ascend_multi_connector.py` | 按子 connector 决定 | 串联多个 connector，并适配 Ascend HMA、多组 KV cache、preempt offload 优先级 | PD + KV Pool、PD + UCM、PD + offload |
| P2P Pull 型 PD | `MooncakeConnectorV1` | `kv_p2p/mooncake_connector.py` | P: `kv_producer`, D: `kv_consumer` | Decode 节点从 Prefill 节点拉取 KV cache | 标准 PD 分离，单机或多机 |
| P2P Layerwise Push 型 PD | `MooncakeLayerwiseConnector` | `kv_p2p/mooncake_layerwise_connector.py` | P: `kv_producer`, D: `kv_consumer` | Prefill 按层推送 KV 到 Decode，计算和传输重叠 | 追求更低 TTFT 的 PD 分离，需要 layerwise proxy |
| Hybrid Attention PD | `MooncakeHybridConnector` | `kv_p2p/mooncake_hybrid_connector.py` | P: `kv_producer`, D: `kv_consumer` | 面向 hybrid attention、多 KV cache group、压缩 KV block 的 PD KV 传输 | DeepSeek V4 Flash/Pro 等 hybrid attention 模型 |
| KV Pool | `AscendStoreConnector`, `MooncakeConnectorStoreV1` | `kv_pool/ascend_store/ascend_store_connector.py` | `kv_both` 或 P/D | 通过 Mooncake/Memcache/Yuanrong 后端扩展 prefix cache 容量 | PD-Mixed、PD + KV Pool、跨实例前缀复用 |
| UCM 存储 | `UCMConnector` | `kv_pool/ucm_connector.py` | `kv_both` 或组合使用 | 委托 UCM 做持久化/分层 KV cache 管理 | 中心化 PD、UCM prefix cache、UCM + Mooncake P2P |
| LMCache-Ascend | `LMCacheAscendConnector` | `kv_pool/lmcache_ascend_connector.py` | `kv_both` | 使用 LMCache-Ascend 的动态 KV Connector | 已采用 LMCache-Ascend 生态的部署 |
| CPU Offload | `SimpleCPUOffloadConnector` | `kv_pool/simple_cpu_offload/simple_cpu_offload_connector.py` | `kv_both` | 覆盖上游 Simple CPU offload，worker 改为 NPU 原生异步拷贝 | NPU cache 不足时用主机内存扩容 |
| Recompute Offload | `RecomputeCPUOffloadConnector` | `kv_pool/recompute_cpu_offload/recompute_cpu_offload_connector.py` | `kv_both` | 对 recompute/preempt 的请求保留 CPU KV，恢复时异步加载 | Decode 侧启用 recompute scheduler 的场景 |

## 通用 NPU 前置条件

### Mooncake P2P 和 Mooncake Store

Mooncake 相关 connector 需要 Ascend Direct 能力。典型安装方式是在 Mooncake 编译时启用：

```bash
cmake .. -DUSE_ASCEND_DIRECT=ON
make -j
make install
export LD_LIBRARY_PATH=/usr/local/lib64/python3.12/site-packages/mooncake:$LD_LIBRARY_PATH
```

KV Pool 使用 Mooncake 后端时，需要 `MOONCAKE_CONFIG_PATH` 指向 `mooncake.json`。NPU 上关键配置如下：

```json
{
  "metadata_server": "P2PHANDSHAKE",
  "protocol": "ascend",
  "device_name": "",
  "master_server_address": "x.x.x.x:50088",
  "global_segment_size": "1GB",
  "preferred_segment": false,
  "prefer_alloc_in_same_node": true
}
```

其中 `protocol` 必须为 `ascend`，`global_segment_size` 在 A3 + FabricMem 场景必须按 1 GB 对齐。启动 Mooncake master 示例：

```bash
mooncake_master --port 50088 \
  --eviction_high_watermark_ratio 0.9 \
  --eviction_ratio 0.1 \
  --default_kv_lease_ttl 11000
```

### 网络和端口

PD 分离依赖 NPU 网络连通性。部署前建议检查：

- A2 常用 `hccn_tool -i <idx> -ip -g` 获取 NPU IP，使用 `hccn_tool -i <idx> -ping -g address <peer-npu-ip>` 验证连通性。
- A3 常用 `hccn_tool -i <idx> -vnic -g` 获取虚拟 NPU IP，使用 `hccn_tool -i <idx> -hccs_ping -g address <peer-vnic-ip>` 验证连通性。
- `GLOO_SOCKET_IFNAME`、`TP_SOCKET_IFNAME`、`HCCL_SOCKET_IFNAME` 需要设置到可达网卡。
- `kv_port`、OpenAI API `--port`、proxy 端口、Mooncake master 端口不要冲突。

## `MooncakeConnectorV1`: 标准 P2P Pull 型 PD

### 功能

`MooncakeConnectorV1` 是标准 PD 分离 connector。Prefill 节点负责计算 prompt 的 KV cache，Decode 节点等待并从 Prefill 节点拉取 KV cache 后继续生成 token。

典型流程：

1. Proxy 将请求发送到 Prefill，携带 `do_remote_decode=True`。
2. Prefill 完成 prefill 后返回远端 KV 元信息，如 block id、host、port、engine id。
3. Proxy 将请求转发给 Decode，携带 `do_remote_prefill=True`。
4. Decode 预分配本地 KV block，通过 Mooncake TransferEngine 拉取远端 KV cache。

### NPU 用法

Prefill 节点：

```bash
vllm serve /path/to/model \
  --host 0.0.0.0 \
  --port 8100 \
  --tensor-parallel-size 2 \
  --data-parallel-size 2 \
  --gpu-memory-utilization 0.8 \
  --kv-transfer-config '{
    "kv_connector": "MooncakeConnectorV1",
    "kv_buffer_device": "npu",
    "kv_role": "kv_producer",
    "kv_parallel_size": 1,
    "kv_port": "20001",
    "kv_rank": 0,
    "kv_connector_extra_config": {
      "prefill": {"dp_size": 2, "tp_size": 2},
      "decode": {"dp_size": 2, "tp_size": 2}
    }
  }'
```

Decode 节点：

```bash
vllm serve /path/to/model \
  --host 0.0.0.0 \
  --port 8200 \
  --tensor-parallel-size 2 \
  --data-parallel-size 2 \
  --gpu-memory-utilization 0.8 \
  --compilation-config '{"cudagraph_mode":"FULL_DECODE_ONLY"}' \
  --kv-transfer-config '{
    "kv_connector": "MooncakeConnectorV1",
    "kv_buffer_device": "npu",
    "kv_role": "kv_consumer",
    "kv_parallel_size": 1,
    "kv_port": "20002",
    "kv_rank": 1,
    "kv_connector_extra_config": {
      "prefill": {"dp_size": 2, "tp_size": 2},
      "decode": {"dp_size": 2, "tp_size": 2}
    }
  }'
```

Python 离线用法同样通过 `KVTransferConfig`：

```python
from vllm.config import KVTransferConfig

prefill_config = KVTransferConfig(
    kv_connector="MooncakeConnectorV1",
    kv_role="kv_producer",
    kv_port="30000",
    engine_id="0",
    kv_connector_extra_config={
        "prefill": {"dp_size": 1, "tp_size": 1},
        "decode": {"dp_size": 1, "tp_size": 1},
    },
)
```

## `MooncakeLayerwiseConnector`: Layerwise Push 型 PD

### 功能

`MooncakeLayerwiseConnector` 把 KV cache 传输拆到 layer 粒度。请求先进入 Decode，Decode 通过 proxy 的 `/v1/metaserver` 反向触发 Prefill；Prefill 在前向过程中逐层保存并推送 KV，Decode 逐层等待加载。这样可以把 Prefill 计算和 KV 传输重叠，降低 TTFT。

vLLM-Ascend 的注意力实现通过以下两个钩子与 connector 协作：

- `wait_for_kv_layer_from_connector(layer_name)`: Decode 在使用某层 KV 前等待该层加载完成。
- `maybe_save_kv_layer_to_connector(layer_name, kv_cache_layer)`: Prefill 在某层 KV 生成后提交给 connector。

### NPU 用法

Prefill 和 Decode 的 `kv_connector_extra_config` 仍需写明 P/D 两侧并行度：

```json
{
  "kv_connector": "MooncakeLayerwiseConnector",
  "kv_role": "kv_producer",
  "kv_port": "30000",
  "kv_connector_extra_config": {
    "prefill": {"dp_size": 2, "tp_size": 8},
    "decode": {"dp_size": 4, "tp_size": 4}
  }
}
```

Decode 侧改为：

```json
{
  "kv_connector": "MooncakeLayerwiseConnector",
  "kv_role": "kv_consumer",
  "kv_port": "30200",
  "kv_connector_extra_config": {
    "prefill": {"dp_size": 2, "tp_size": 8},
    "decode": {"dp_size": 4, "tp_size": 4}
  }
}
```

必须使用 layerwise proxy，例如：

```bash
python examples/disaggregated_prefill_v1/load_balance_proxy_layerwise_server_example.py \
  --host <proxy-ip> \
  --port <proxy-port> \
  --prefiller-hosts <p-host-1> <p-host-2> \
  --prefiller-ports <p-port-1> <p-port-2> \
  --decoder-hosts <d-host-1> <d-host-2> \
  --decoder-ports <d-port-1> <d-port-2>
```

`--host` 需要对 Decode 节点可达，因为 Decode 会回连 `http://<host>:<port>/v1/metaserver`。

## `MooncakeHybridConnector`: Hybrid Attention 模型的 PD 连接器

### 功能

`MooncakeHybridConnector` 面向 hybrid attention、多 KV cache group、sliding window 或压缩 KV cache 的模型。DeepSeek V4 Flash/Pro 教程中使用它做 PD KV 传输。

与 `MooncakeConnectorV1` 相比，它额外处理：

- 多个 KV cache group 的 HMA 元数据。
- hybrid attention 下的不同 block size、sliding window、compress ratio。
- Prefill/Decode TP、DP、PP 映射。

### 当前限制

从实现看，`MooncakeHybridConnector` 有几个硬限制：

- `prefill_context_parallel_size * decode_context_parallel_size == 1`，也就是当前不支持 CP world size 大于 1。
- Prefill TP size 需要大于等于 Decode TP size。
- Decode PP size 必须为 1。

### NPU 用法

配置形态与 Mooncake PD 类似，只是 connector 名称改为 `MooncakeHybridConnector`：

```json
{
  "kv_connector": "MooncakeHybridConnector",
  "kv_role": "kv_producer",
  "kv_port": "30000",
  "kv_connector_extra_config": {
    "prefill": {"dp_size": 1, "tp_size": 8},
    "decode": {"dp_size": 1, "tp_size": 8}
  }
}
```

Decode 侧使用 `kv_role: "kv_consumer"`，并保持 `prefill` 和 `decode` 的并行配置与实际服务参数一致。对 DeepSeek V4 类 PD 场景，Decode 侧通常建议启用 recompute scheduler，使 Decode KV 不足时可以回到 Prefill 侧重算。

## `AscendStoreConnector`: KV Pool 和后端存储

### 功能

`AscendStoreConnector` 是 vLLM-Ascend 的 KV Pool connector。它不是直接的 P/D P2P 传输通道，而是把 prefix KV cache 存入外部池，并在后续请求中按 prefix 命中复用。

当前后端包括：

- `backend: "mooncake"`: Mooncake Store，NPU 上要求 `protocol: "ascend"`。
- `backend: "memcache"`: Memcache 后端，支持 layerwise KV Pool。
- `backend: "yuanrong"`: Yuanrong Datasystem 后端。

`MooncakeConnectorStoreV1` 是同一实现的兼容注册名，推荐新配置直接使用 `AscendStoreConnector`。

### PD-Mixed 用法

单个实例既生产也消费 KV cache 时使用 `kv_both`：

```bash
export MOONCAKE_CONFIG_PATH=/path/to/mooncake.json

vllm serve /path/to/model \
  --host 0.0.0.0 \
  --port 8100 \
  --block-size 128 \
  --max-num-batched-tokens 16384 \
  --kv-transfer-config '{
    "kv_connector": "AscendStoreConnector",
    "kv_role": "kv_both",
    "kv_load_failure_policy": "recompute",
    "kv_connector_extra_config": {
      "lookup_rpc_port": "1",
      "backend": "mooncake"
    }
  }'
```

### PD 分离 + KV Pool

PD 分离里常用 `MultiConnector` 同时启用：

- `MooncakeConnectorV1` 负责 P/D 间 KV transfer。
- `AscendStoreConnector` 负责 Prefill 侧 prefix cache pool。

Prefill 节点示例：

```json
{
  "kv_connector": "MultiConnector",
  "kv_role": "kv_producer",
  "kv_load_failure_policy": "recompute",
  "kv_connector_extra_config": {
    "connectors": [
      {
        "kv_connector": "MooncakeConnectorV1",
        "kv_role": "kv_producer",
        "kv_port": "20001",
        "kv_connector_extra_config": {
          "prefill": {"dp_size": 1, "tp_size": 1},
          "decode": {"dp_size": 1, "tp_size": 1}
        }
      },
      {
        "kv_connector": "AscendStoreConnector",
        "kv_role": "kv_producer",
        "kv_connector_extra_config": {
          "lookup_rpc_port": "0",
          "backend": "mooncake"
        }
      }
    ]
  }
}
```

Decode 节点同理把两个子 connector 的 `kv_role` 改为 `kv_consumer`。默认情况下，KV Pool 的 put/get 主要发生在 Prefill 节点；MLA 场景若需要 Decode 节点写入给 Prefill 复用，可以给 `AscendStoreConnector` 增加 `consumer_is_to_put: true`。

### SSD Offload 和 FabricMem

Mooncake Store 支持 SSD offload 时，在 `mooncake.json` 增加：

```json
{
  "enable_ssd_offload": true,
  "ssd_offload_path": "/nvme/mooncake_offload"
}
```

同时启动 master 时加 `--enable_offload=true`。A3 + `ASCEND_ENABLE_USE_FABRIC_MEM=1` 时，`global_segment_size` 和 `MOONCAKE_OFFLOAD_LOCAL_BUFFER_SIZE_BYTES` 都必须按 1 GB 对齐，并计入每 rank 的 fabric memory 预算。

## Layerwise KV Pool: `AscendStoreConnector` 的逐层模式

Layerwise KV Pool 和 `MooncakeLayerwiseConnector` 不是同一件事。前者是 `AscendStoreConnector` 的 KV Pool 优化，后者是 P2P PD 连接器。

开启方式是在 `AscendStoreConnector` 的 extra config 中增加：

```json
{
  "kv_connector": "AscendStoreConnector",
  "kv_role": "kv_both",
  "kv_connector_extra_config": {
    "backend": "memcache",
    "use_layerwise": true,
    "layerwise_prefetch_layers": 1,
    "layerwise_max_transfer_blocks": 0,
    "layerwise_max_transfer_bytes": 0,
    "h2d_stagger_us": 0
  }
}
```

当前限制：

- 逐层 KV Pool 目前只支持 `backend: "memcache"`。
- 不支持 hybrid KV cache groups。
- 尚未与 CP attention backend 集成。
- PD 分离场景需要 layerwise proxy，因为标准 disagg proxy 不提供 `/v1/metaserver`。

## `UCMConnector`: UCM 统一缓存管理

### 功能

`UCMConnector` 是对外部 UCM connector 的封装，提供三层缓存结构：

```text
HBM/NPU memory -> local DRAM cache -> storage backend
```

UCM 能力包括 prefix cache、sparse attention、中心化 PD，以及 UCM + Mooncake 的 P2P PD。

### 中心化 PD 用法

Prefill 和 Decode 都使用 `UCMConnector`：

```json
{
  "kv_connector": "UCMConnector",
  "kv_role": "kv_both",
  "kv_connector_extra_config": {
    "UCM_CONFIG_FILE": "/path/to/ucm_config_example.yaml"
  }
}
```

### UCM + Mooncake P2P 用法

Prefill 节点用 `MultiConnector` 组合 Mooncake P2P 和 UCM prefix cache：

```json
{
  "kv_connector": "MultiConnector",
  "kv_role": "kv_producer",
  "kv_connector_extra_config": {
    "connectors": [
      {
        "kv_connector": "MooncakeConnectorV1",
        "kv_role": "kv_producer",
        "kv_port": 20001,
        "kv_connector_extra_config": {
          "prefill": {"dp_size": 1, "tp_size": 4},
          "decode": {"dp_size": 1, "tp_size": 4}
        }
      },
      {
        "kv_connector": "UCMConnector",
        "kv_role": "kv_both",
        "kv_connector_extra_config": {
          "UCM_CONFIG_FILE": "/path/to/ucm_config_example.yaml"
        }
      }
    ]
  }
}
```

Decode 节点一般只配置 `MooncakeConnectorV1` 消费 Prefill 传来的 KV。

## `LMCacheAscendConnector`

`LMCacheAscendConnector` 依赖 LMCache-Ascend 包。vLLM-Ascend 侧实现很薄，只导入 `lmcache_ascend` 并复用 vLLM 的 `LMCacheConnectorV1`。

在线服务用法：

```bash
python -m vllm.entrypoints.openai.api_server \
  --port 8100 \
  --model /data/models/Qwen/Qwen3-32B \
  --trust-remote-code \
  --block-size 128 \
  --kv-transfer-config '{"kv_connector":"LMCacheAscendConnector","kv_role":"kv_both"}'
```

离线用法：

```python
from vllm.config import KVTransferConfig

ktc = KVTransferConfig(
    kv_connector="LMCacheAscendConnector",
    kv_role="kv_both",
)
```

## CPU Offload 相关用法

### 上游 `OffloadingConnector` + `NPUOffloadingSpec`

Feature guide 中的 CPU offload 推荐配置基于 vLLM 的 `OffloadingConnector`，并通过 `NPUOffloadingSpec` 把数据搬运适配到 NPU：

```json
{
  "kv_connector": "OffloadingConnector",
  "kv_role": "kv_both",
  "kv_connector_extra_config": {
    "num_cpu_blocks": 1000,
    "block_size": 128,
    "spec_name": "NPUOffloadingSpec",
    "spec_module_path": "vllm_ascend.kv_offload.npu"
  }
}
```

其作用是把非活跃 KV block 从 NPU 异步拷到 CPU，后续 prefix cache 在 NPU miss 但 CPU hit 时再通过 H2D NPU stream 加载回来。

### `SimpleCPUOffloadConnector`

vLLM-Ascend 覆盖上游 `SimpleCPUOffloadConnector`，worker 替换为 NPU 版本，使用 `aclrtMemcpyBatchAsync` 和 `torch.npu` stream/event。

典型配置：

```python
from vllm.config import KVTransferConfig

kv_transfer_config = KVTransferConfig(
    kv_connector="SimpleCPUOffloadConnector",
    kv_role="kv_both",
    kv_connector_extra_config={
        "cpu_bytes_to_use": 1 << 30,
        "lazy_offload": False,
    },
)
```

需要开启 prefix caching 才能发挥 CPU 侧命中价值。

### `RecomputeCPUOffloadConnector`

`RecomputeCPUOffloadConnector` 面向 recompute/preempt 场景。它在请求被抢占并需要重算时，把可保留的 KV cache offload 到 CPU，后续恢复时加载回 NPU。关键参数：

```json
{
  "kv_connector": "RecomputeCPUOffloadConnector",
  "kv_role": "kv_both",
  "kv_connector_extra_config": {
    "cpu_bytes_to_use": 8589934592,
    "enable_offload_prefix_caching": false
  }
}
```

如果需要明确每个 rank 的预算，也可以设置 `cpu_bytes_to_use_per_rank`。默认总 CPU 容量为 8 GiB，并按 world size 切分。

旧的 `CPUOffloadingConnector` 代码仍可在源码中看到，但 release notes 已标记为 deprecated，新部署不要优先选择它。

## 选择建议

| 目标 | 推荐 connector |
| --- | --- |
| 标准 PD 分离，P 计算、D 拉取 KV | `MooncakeConnectorV1` |
| 希望 Prefill 计算和 KV 传输按层重叠 | `MooncakeLayerwiseConnector` + layerwise proxy |
| DeepSeek V4/Qwen hybrid attention 这类多 KV cache group 模型做 PD | `MooncakeHybridConnector` |
| 扩大 prefix cache 容量，跨请求/跨实例复用 | `AscendStoreConnector` |
| PD + KV Pool | `MultiConnector` + `MooncakeConnectorV1` + `AscendStoreConnector` |
| 使用 UCM 的持久化/中心化缓存 | `UCMConnector` |
| 使用 LMCache-Ascend 生态 | `LMCacheAscendConnector` |
| 用主机内存扩展 NPU KV cache | `OffloadingConnector` + `NPUOffloadingSpec` 或 `SimpleCPUOffloadConnector` |
| Recompute/preempt 请求希望保留 CPU KV | `RecomputeCPUOffloadConnector` |

## 参考来源

- 本地实现注册表：`vllm_ascend/distributed/kv_transfer/__init__.py`
- PD 设计文档：`docs/source/developer_guide/Design_Documents/disaggregated_prefill.md`
- KV Pool 设计文档：`docs/source/developer_guide/Design_Documents/KV_Cache_Pool_Guide.md`
- Feature Guide 入口：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/index.html>
- KV Pool Guide：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/kv_pool.html>
- KV Cache CPU Offload Guide：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/kv_cache_cpu_offload.html>
- UCM Store Deployment Guide：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/ucm_deployment.html>
- LMCache-Ascend Deployment Guide：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/user_guide/feature_guide/lmcache_ascend_deployment.html>
- PD Mooncake 多机教程：<https://docs.vllm.ai/projects/ascend/zh-cn/latest/tutorials/features/pd_disaggregation_mooncake_multi_node.html>
