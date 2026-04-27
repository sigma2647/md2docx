#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PROJECT="md2docx"
CONFIG="Release"
RID="linux-x64"
PUBLISH_DIR="bin/${CONFIG}/net10.0/${RID}/publish"
INSTALL_DIR="${HOME}/.local/bin"

usage() {
    echo "Usage: $0 [build|install|clean]"
    echo ""
    echo "  build   - 快速编译（framework-dependent，开发用）"
    echo "  install - 单文件自包含发布并安装到 ~/.local/bin"
    echo "  clean   - 清理构建产物"
    exit 1
}

cmd="${1:-install}"

case "$cmd" in
    build)
        echo "=> Building (framework-dependent)..."
        dotnet build -c "$CONFIG"
        echo "Done. Output: bin/${CONFIG}/net10.0/${PROJECT}"
        ;;

    install)
        echo "=> Publishing single-file self-contained binary..."
        dotnet publish -c "$CONFIG" \
            -r "$RID" \
            --self-contained true \
            -p:PublishSingleFile=true \
            -p:PublishTrimmed=false

        echo "=> Installing to ${INSTALL_DIR}..."
        mkdir -p "$INSTALL_DIR"
        cp -f "${PUBLISH_DIR}/${PROJECT}" "${INSTALL_DIR}/${PROJECT}"
        chmod +x "${INSTALL_DIR}/${PROJECT}"

        echo "Done. Installed: ${INSTALL_DIR}/${PROJECT}"
        echo "Run '${PROJECT} --help' to verify."
        ;;

    clean)
        echo "=> Cleaning..."
        rm -rf bin/ obj/
        echo "Done."
        ;;

    *)
        usage
        ;;
esac
