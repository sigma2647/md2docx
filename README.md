
ParseMarkdown

## 表格使用方法
```
| 姓名 | 年龄 | 职位   |
|------|------|--------|
| 张三 | 28   | 工程师 |
| 李四 | 35   | 经理   |
[各部门人员统计]

| 指标 | 数值 |
|------|------|
| 总数 | 100  |
[无标题时不写此行，也会自动添加 表 N]
```
## 字体对应



## 题注

字体
```
10
黑体
```

| 目标字号      | OpenXML Val |
|---------------|-------------|
| 二号 (22pt)   | `"44"`      |
| 三号 (16pt)   | `"32"`      |
| 小四 (12pt)   | `"24"`      |
| 五号 (10.5pt) | `"21"`      |
| 小五 (9pt)    | `"18"`      |



| 目标字号 | 磅值 | OpenXML Val |
|----------|------|-------------|
| 二号     | 22   | `"44"`      |
| 小二     | 18   | `"36"`      |
| 三号     | 16   | `"32"`      |
| 小三     | 15   | `"30"`      |
| 四号     | 14   | `"28"`      |
| 小四     | 12   | `"24"`      |
| 五号     | 10.5 | `"21"`      |
| 10pt     | 10   | `"20"`      |
| 小五     | 9    | `"18"`      |
| 六号     | 7.5  | `"15"`      |
| 小六     | 6.5  | `"13"`      |
| 七号     | 5.5  | `"11"`      |
| 八号     | 5    | `"10"`      |



## Build & Install

### 方法 1：Makefile（推荐）

```bash
# 编译 + 安装到 ~/.local/bin
make install

# 仅快速编译（framework-dependent，开发调试用）
make build

# 清理构建产物
make clean
```

### 方法 2：build.sh

```bash
# 编译 + 安装到 ~/.local/bin
./build.sh install

# 仅快速编译
./build.sh build

# 清理
./build.sh clean
```

### 方法 3：手动命令

```bash
# 单文件自包含发布（产生唯一可执行文件，零依赖）
dotnet publish -c Release -r linux-x64 \
  --self-contained true \
  -p:PublishSingleFile=true

# 复制到 PATH
cp bin/Release/net10.0/linux-x64/publish/md2docx ~/.local/bin/
chmod +x ~/.local/bin/md2docx
```

> ⚠️ **注意**：不要对 framework-dependent 的发布结果使用 `ln -s` 软链接到 PATH。运行时会在链接所在目录（如 `~/.local/bin/`）查找 DLL，导致文件丢失错误。单文件自包含（`-p:PublishSingleFile=true`）发布结果可直接复制或软链接到任意位置。

