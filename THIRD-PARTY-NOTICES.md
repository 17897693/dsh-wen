# 第三方组件声明（THIRD-PARTY NOTICES）

本包内含的第三方二进制与模型**不属于**本项目的 MIT 许可范围，各自遵循其原始许可。
若你再次分发本包，请连同本文件一并保留。

---

## 1. RapidOCR-json（OCR 引擎）

- **文件**：`vendor/RapidOCR-json_v0.2.0/RapidOCR-json.exe`
- **上游**：<https://github.com/hiroi-sora/RapidOCR-json>
- **描述**：OCR 离线图片文字识别命令行 Windows 程序，以 JSON 字符串形式输出结果。
- **许可证**：MIT

该程序的 MIT 许可要求再分发时保留其版权声明与许可声明。上游 LICENSE 全文见
<https://github.com/hiroi-sora/RapidOCR-json/blob/main/LICENSE>
（版权归 hiroi-sora 及 RapidOCR-json 项目贡献者；若你需要逐字准确的副本，
请从上述地址下载 LICENSE 替换本节说明）。

```
MIT License

Copyright (c) hiroi-sora and RapidOCR-json contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 2. PaddleOCR PP-OCRv4 模型

- **文件**：`vendor/RapidOCR-json_v0.2.0/models/` 下的
  `ch_PP-OCRv4_det_infer.onnx`、`ch_PP-OCRv4_rec_infer.onnx`、
  `ch_ppocr_mobile_v2.0_cls_infer.onnx`、`dict_chinese.txt`
- **上游**：<https://github.com/PaddlePaddle/PaddleOCR>
- **许可证**：Apache License 2.0（以 PaddleOCR 上游声明为准）

## 3. 说明

- 上述组件由**上游项目**提供，本项目仅做集成与调用，不对其做任何修改或再许可。
- 如果你只需要本项目自研的解析能力（docx / xlsx / pptx / pdf / odf / csv / html / md），
  完全可以不打包 `vendor/`，此时本声明中第 1、2 节不适用于你的分发包。
