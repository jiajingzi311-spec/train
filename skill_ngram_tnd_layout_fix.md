---
name: ngram-tnd-layout-fix
description: vllm-ascend ngram speculative decoding + FULL cudagraph + TND layout 崩溃修复方案
type: reference
originSessionId: 3af7a370-2df4-4451-a047-0b22d74f9fa7
---
# Skill: 修复 ngram + FULL cudagraph + TND layout 崩溃

## 问题描述

**现象**：启用 ngram speculative decoding + FULL cudagraph 时，NPU FA 算子报错：
```
RuntimeError: npu_fused_infer_attention_score_out_symint NPU function error
When layout is TND, queryT(16) must be equal to the last element of actualSequenceLengthQ(13)
```

**根因**：`_pad_query_start_loc_for_fia()` 中 uniform-batch 判断条件不充分。

具体场景（ngram spec_token_num=3）：
- `uniform_decode_query_len = 1 + 3 = 4`
- 4 个请求实际 token 分布为 `[4, 4, 4, 1]`，`query_start_loc = [0, 4, 8, 12, 13]`
- padding 后 `num_tokens_padded = 16`，`num_reqs_padded = 4`
- 原条件 `16 == 4 * 4` 为 True → 进入 uniform-batch 分支
- 但 uniform 分支不修正 `query_start_loc[num_reqs]`（仍为 13），导致 `actual_seq_lengths_q[-1] = 13 ≠ queryT = 16`
- NPU FA 算子检查 TND 约束失败，崩溃

**为什么只有 ngram + FULL cudagraph 才触发**：
- 无 spec decoding时 `uniform_decode_query_len = 1`，每个 decode 请求恰好 1 token，天然 uniform
- ngram 使 `uniform_decode_query_len = 4`，但部分请求可能只有 1 token（如刚从 prefill 转入 decode 的请求），造成实际不 uniform
- FULL cudagraph 模式下 `num_reqs_padded = num_reqs`（不做额外 padding），使条件恰好为 True

## 修复方案

### 修改文件

`vllm-ascend/vllm_ascend/worker/model_runner_v1.py`

### 修改位置

`_pad_query_start_loc_for_fia()` 方法，约 line 586

### 修改内容

**原代码**（line 586）：
```python
        if num_tokens_padded == num_reqs_padded * self.uniform_decode_query_len:
```

**改为**：
```python
        if (num_tokens_padded == num_reqs_padded * self.uniform_decode_query_len
                and self.query_start_loc.np[num_reqs] == num_tokens_padded):
```

### 修复逻辑

增加 `self.query_start_loc.np[num_reqs] == num_tokens_padded` 检查，确保不仅是 padding 后总数匹配 uniform 假设，而且实际的 `query_start_loc` 末尾值也匹配。当实际不 uniform 时，走 mixed-batch 分支，插入 dummy request 来满足 TND 约束。

### 修复前后对比

| 场景 | 原条件结果 | 新条件结果 | 走哪个分支 |
|------|-----------|-----------|-----------|
| 真正 uniform（4 请求各 4 token） | True | True（13→16 时 False，但真正 uniform 时 query_start_loc[-1]=16） | uniform ✓ |
| ngram 不 uniform（`[4,4,4,1]`，pad 后 16） | True ✗ | False ✓ | mixed ✓ |
| 非 cudagraph 的 mixed batch | 走 else | 走 else | mixed ✓ |

### 安全保障

1. **备份原文件**：修改前先 `cp model_runner_v1.py model_runner_v1.py.bak`
2. **只改一行条件**，不改任何逻辑分支
3. **回退命令**：
   ```bash
   cp model_runner_v1.py.bak model_runner_v1.py
   ```

## 验证方法

1. 不开 spec decoding 启动 vLLM → 正常运行（回归验证）
2. 开 ngram spec decoding + FULL cudagraph 启动 → 不再报 TND 约束错误
3. 观察日志中 `actual_seq_lengths_q` 末尾值是否等于 `queryT`
4. 配合 Mooncake connector / prefix cache 等功能验证无回归
