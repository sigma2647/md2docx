# md2docx Makefile
# Usage:
#   make build      # 快速编译（framework-dependent，开发用）
#   make install    # 单文件自包含发布并安装到 ~/.local/bin
#   make clean      # 清理构建产物

PROJECT      := md2docx
CONFIG       := Release
RID          := linux-x64
PUBLISH_DIR  := bin/$(CONFIG)/net10.0/$(RID)/publish
INSTALL_DIR  := $(HOME)/.local/bin
EXECUTABLE   := $(PUBLISH_DIR)/$(PROJECT)

.PHONY: all build install clean

all: install

## 快速编译（不打包 Runtime，仅验证）
build:
	dotnet build -c $(CONFIG)

## 单文件自包含发布 + 安装到 PATH
# 这样只产生一个可执行文件，可直接复制/链接到任意位置，不会丢失 DLL
install:
	@echo "=> Publishing single-file self-contained binary..."
	dotnet publish -c $(CONFIG) \
		-r $(RID) \
		--self-contained true \
		-p:PublishSingleFile=true \
		-p:PublishTrimmed=false
	@echo "=> Installing to $(INSTALL_DIR)..."
	@mkdir -p $(INSTALL_DIR)
	@cp -f $(EXECUTABLE) $(INSTALL_DIR)/$(PROJECT)
	@chmod +x $(INSTALL_DIR)/$(PROJECT)
	@echo "Done. Installed: $(INSTALL_DIR)/$(PROJECT)"
	@echo "Run '$(PROJECT) --help' to verify."

## 清理
 clean:
	rm -rf bin/ obj/
