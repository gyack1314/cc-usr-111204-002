# 燃机技术工作簿的单元格来源化抽取

此Python项目管理独立的能源领域合成源资料，`data/`为原始样本及访问配置，`source_data.io`负责严格读取和字节身份，`source_data.domain`提供现有库存摘要，`source_data.replay`按声明顺序返回不解释的源响应。字段和支持的来源口径见[资料约定](docs/data-contract.md)。所有机构、设备、对象与数值均为虚构。

Python 3.11及以上即可使用，无第三方Python依赖。项目根目录运行：

```sh
python -m unittest discover -s tests -v
python -m compileall -q source_data tests tools
python -m source_data.cli inspect data/normal.json
python -m source_data.cli collect data/normal.json out/normal.jsonl
python -I tools/make_samples.py
```

`inspect`展示已有的原始库存，不修改资料。`collect`按配置的sheet名称与行区间对.xlsx做逐格来源化采集，首行输出meta、其后每行一条JSONL记录，保留坐标、表头来源、原类型、样式、原公式与展开式、缓存状态和数值词法/Decimal；口径见[资料约定](docs/data-contract.md)。读取外来资料建议用`python -I`并传路径参数，脚本目录与资料目录分离。`invalid-`前缀文件为不合法输入，其他变体保留不同业务边界；源协议与二进制格式使用资料约定中的有限范围，未解释部分仍以原字节保存。

固定基础镜像见Dockerfile，可执行`docker build -t energy-source .`，再用`docker run --rm energy-source`显示摘要；`docker run --rm energy-source python -m unittest discover -s tests -v`执行已有检查。资料和回放不依赖真实设备、账户或实时网络服务。
