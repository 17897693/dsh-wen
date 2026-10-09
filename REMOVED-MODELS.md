本文件为 CRLF 例外（全 CRLF 行尾），跨侧搬运按字节复制
# 已移除的 OCR 模型（C 方案瘦身）

本插件默认路径只使用 `rapidocr.js` 中 `DEFAULT_MODELS` 的四个文件：
`ch_PP-OCRv4_det_infer.onnx` / `rec_ch_PP-OCRv4_infer.onnx` / `ch_ppocr_mobile_v2.0_cls_infer.onnx` / `dict_chinese.txt`。

以下文件仅能通过环境变量 `DSH_OFFICE_OCR_DET` / `DSH_OFFICE_OCR_REC` / `DSH_OFFICE_OCR_KEYS` 选用，
简体中文资料（考公考编）路径永不加载，已从现役副本与三份备份中一并删除。

删除前后实测：用真实中文扫描件 PDF 渲染 3 页对比，OCR 输出逐字节相同。

| 文件 | 体积 | 用途 |
| --- | --- | --- |
| `rec_chinese_cht_PP-OCRv3_infer.onnx` | 10.65 MB | 繁体识别 |
| `ch_PP-OCRv3_rec_infer.onnx` | 10.2 MB | PP-OCRv3 简体识别（已被 v4 取代） |
| `rec_japan_PP-OCRv3_infer.onnx` | 9.64 MB | 日文识别 |
| `rec_korean_PP-OCRv3_infer.onnx` | 9.46 MB | 韩文识别 |
| `rec_cyrillic_PP-OCRv3_infer.onnx` | 8.57 MB | 西里尔识别 |
| `rec_en_PP-OCRv3_infer.onnx` | 8.56 MB | 英文识别 |
| `ch_PP-OCRv3_det_infer.onnx` | 2.32 MB | PP-OCRv3 检测（已被 v4 取代） |
| `dict_chinese_cht.txt` | 0.03 MB | 繁体字典 |
| `ppocr_keys_v1.txt` | 0.03 MB | 与 dict_chinese.txt 逐字节重复 |
| `dict_japan.txt` | 0.02 MB | 日文字典 |
| `dict_korean.txt` | 0.01 MB | 韩文字典 |
| `dict_en.txt` | 0 MB | 英文字典 |
| `dict_cyrillic.txt` | 0 MB | 西里尔字典 |

合计约 59.5 MB。

## 恢复方法

从 RapidOCR-json v0.2.0 发行包取回同名文件，放回 `vendor/RapidOCR-json_v0.2.0/models/`，
再设置对应环境变量即可。`configs.txt` / `cmd.txt` 仍列出这些模型集，现已过时（保留未删，共约 2.4 KB）。
