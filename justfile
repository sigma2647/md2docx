# md2docx justfile
# Usage:
#   just build      快速编译（framework-dependent，开发用）
#   just install    单文件自包含发布并安装到 ~/.local/bin
#   just clean      清理构建产物

project     := "md2docx"
config      := "Release"
install_dir := home_dir() / ".local/bin"

# 平台自动检测 → .NET RID
arch := if arch() == "x86_64" { "x64" } else if arch() == "aarch64" { "arm64" } else { arch() }
rid  := if os() == "linux" { "linux-" + arch } else if os() == "macos" { "osx-" + arch } else { "win-" + arch }

publish_dir := "bin" / config / "net10.0" / rid / "publish"
executable  := publish_dir / project

default: install

# 快速编译（不打包 Runtime，仅验证）
build:
    dotnet build -c {{config}}

# 单文件自包含发布 + 安装到 PATH
install:
    @echo "=> Publishing single-file self-contained binary..."
    dotnet publish -c {{config}} \
        -r {{rid}} \
        --self-contained true \
        -p:PublishSingleFile=true \
        -p:PublishTrimmed=false
    @echo "=> Installing to {{install_dir}}..."
    @mkdir -p {{install_dir}}
    @cp -f {{executable}} {{install_dir}}/{{project}}
    @chmod +x {{install_dir}}/{{project}}
    @echo "Done. Installed: {{install_dir}}/{{project}}"
    @echo "Run '{{project}} --help' to verify."

# 清理构建产物
clean:
    rm -rf bin/ obj/
