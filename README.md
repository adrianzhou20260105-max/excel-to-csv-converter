# SheetCast 🚀

> **Excel 批量转 CSV 在线工具** — 100% 本地浏览器解析，文件不经过任何服务器，极致隐私保护，多 Sheet 一键打包 ZIP 下载！

---

## 🌟 项目亮点

*   **🔒 100% 数据隐私安全**：完全在客户端使用浏览器运行，文件不需要上传到任何服务器，企业敏感财务、数据绝对保密。
*   **📂 多 Sheet 批量转换**：自动解析 Excel（支持 `.xlsx`, `.xls`, `.xlsm`, `.xlsb`）中的所有工作表（Sheets），可一键打包 ZIP 压缩包下载，也可单独下载某一个 Sheet。
*   **💎 防止数字精度丢失**：防止大数（如身份证号、手机号、银行卡号等）因自动类型转换丢失前导 `0` 或被强行转换为科学计数法（如 `1.23E+17`）。
*   **🇨🇳 原生 UTF-8 With BOM 编码**：专为中文及跨平台 Excel 设计，生成的 CSV 能够直接在微软 Excel 中双击打开而不产生乱码。
*   **⚡ 极简轻量级，零构建门槛**：基于原生 HTML5 + CSS3 + Javascript (ES Modules) 并通过 CDN 引入依赖，无需搭建复杂的 `Node.js` 构建环境，克隆项目后双击 `index.html` 即可在本地完美运行！
*   **🌗 现代清新双色主题**：设计精美的用户界面，内置光感自适应与手动切换的深浅色模式，采用卡片式精细排版和微交互动画。

---

## 📸 界面预览

*   **拖拽上传区**：直接拖拽文件或点击选择。
*   **文件详情及全局操作**：显示解析到的文件名、大小、Sheet 总数；提供一键打包下载所有 CSV (ZIP)。
*   **Sheet 选择列表**：列出 Excel 中的所有 Sheet 名称，并显示各自的行列数 (`Row x Col`)。
*   **Sheet 数据预览**：展示选中 Sheet 的前 15 行表格视图，列头采用 Excel 标准的 `A`, `B`, `C` 字母编码，单元格内容过长支持悬停浮动查看。

---

## 🛠️ 技术栈

*   **样式库**：[Tailwind CSS (Play CDN)](https://tailwindcss.com/) — 快速构建现代化的响应式界面。
*   **Excel 解析引擎**：[SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) — 强悍的客户端 Excel 解析库。
*   **压缩包生成器**：[JSZip](https://stuk.github.io/jszip/) — 纯 JS 实现的客户端 ZIP 文件打包工具。
*   **图标库**：[FontAwesome 6](https://fontawesome.com/) — 精美矢量功能图标。

---

## 🚀 快速开始

### 方式 1：双击本地运行
1.  下载或克隆本项目到本地。
2.  双击 `index.html` 文件即可在浏览器中直接打开使用。

### 方式 2：使用本地服务运行（推荐）
如果你需要在本地启用完整的模块化服务，可以在项目根目录下执行以下任一命令：

*   **Python**:
    ```bash
    python -m http.server 8000
    ```
    然后在浏览器中打开 `http://localhost:8000`

*   **NodeJS (npx)**:
    ```bash
    npx serve
    ```

---

## 🌍 部署到 GitHub Pages (发布在网上)

由于项目完全是静态的（无后端），你可以非常轻松地将其免费部署到 GitHub Pages 上，供大家或自己在线访问。

1.  在 GitHub 上新建一个 Repository (例如命名为 `excel-to-csv-converter`)。
2.  将本地的所有文件（`index.html`, `app.js`, `style.css`, `README.md`, `LICENSE`）提交并推送（Push）到该仓库的 `main` 分支。
3.  打开你的 GitHub 仓库页面，点击右侧的 **Settings** (设置)。
4.  在左侧菜单栏中选择 **Pages**。
5.  在 **Build and deployment** 下的 **Source** 选择 `Deploy from a branch`。
6.  在 **Branch** 下拉菜单中选择 `main` 分支和 `/ (root)` 根目录，点击 **Save** (保存)。
7.  稍等 1-2 分钟，GitHub Pages 将会为您生成一个在线链接（格式为 `https://<your-username>.github.io/excel-to-csv-converter/`），打开即可在线使用拖拽转换！

---

## ⚖️ 开源协议

本项目基于 **[MIT License](LICENSE)** 协议开源，欢迎任意克隆、修改、分发或在您的商业项目中使用。
