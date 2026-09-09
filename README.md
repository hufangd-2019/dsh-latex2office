# dsh-latex2office

LaTeX → Office 公式插件（[DeepSeek Harness](https://www.deepseek.com) / DSH Desktop 永久本地工具）。

把 LaTeX 数学公式转成 **Word/WPS 原生可编辑 OMML 公式**或高清渲染图片，直接插入 `.docx` / `.pptx`，支持批量原子插入、学术制表位编号（公式居中 + 编号右缘）、自动备份与回读验证。

## 注册的工具

| 工具 | 作用 |
|---|---|
| `latex2office_insert` | 把 LaTeX 公式插入**现有** .docx / .pptx（OMML 原生公式或渲染图片） |
| `latex2office_generate` | 从公式列表**生成新** .docx |
| `latex2office_preview` | 渲染单条公式为 PNG 预览图（插入前先目检） |

行为要点：

- **docx 默认 `math_mode=omml`**：生成原生 OMML 方程，Word / WPS 文字中可双击进入公式编辑器继续编辑；pptx 默认 `math_mode=image`（WPS 演示对 pptx 内嵌 OMML 有已知的不显示风险），可按条传 `omml` 覆盖
- 插入位置：`end`（默认）/ `after_text` / `before_text` / `after_paragraph`；docx 支持 `display:false` 行内插入（公式嵌进文字流）
- **批量原子**：任何一条失败（坏 LaTeX、找不到锚点），整个文件一字节不动；锚点一律按原始文档内容匹配
- **编号**：每条公式可带 `number:"3.1"` → 学术制表位排版（公式居中 + `(3.1)` 贴右页边）
- 首次修改自动生成 `原名.bak.docx` / `原名.bak.pptx`（已存在则永不覆盖）；每次写入后自动回读验证公式数量
- 复杂结构自动映射：`aligned/align`→`m:eqArr`、`cases`/`\left(...\right)`→`m:d`、矩阵环境→`m:m`

## 宿主依赖（非 npm，需本机安装）

| 依赖 | 用途 | 说明 |
|---|---|---|
| Python ≥ 3.10 | 文档引擎 | 需 `python-docx`、`python-pptx`、`lxml`、`Pillow`（`matplotlib` 可选降级） |
| pandoc | LaTeX→OMML | 经临时 docx 提取 `m:oMathPara` 拼接进目标文档 |
| LibreOffice | 图片渲染 | `soffice --headless`，使用独立临时 profile，不与用户实例互锁 |

解析顺序：cordis 行配置 → PATH → 常见安装路径表。任一缺失时**只让受影响的工具返回清晰 JSON 错误**，不影响插件加载与 DSH 运行。

## 安装

### 方式一：DSH 桌面端插件市场

在 dshmarket 中搜索 `dsh-latex2office` 一键安装，重启 DSH 桌面端生效。

### 方式二：手动（GitHub 依赖）

编辑 DSH 数据目录下的 profile manifest（桌面端为 `harness/profiles/web/package.json`）：

```jsonc
{
  "dependencies": {
    "dsh-latex2office": "github:hufangd-2019/dsh-latex2office"   // ① 加依赖
  },
  "dsh": {
    "profile": {
      "bundles": [
        "dsh-latex2office"              // ② 加 bundle（顺序即 patch 应用顺序）
      ]
    }
  }
}
```

然后在 profile 目录执行 `pnpm install`，重启 DSH 桌面端（bundle 层在启动时解析）。

### 方式三：本地开发（link）

`dependencies` 写 `"dsh-latex2office": "link:<本地目录>"`，其余同上。改代码后重启即生效，适合二次开发。

### 可选配置覆盖

在 profile 自身的 `cordis.patch.yml` 中给本插件的行挂 config：

```yaml
- id: latex2office-tools
  config:
    pythonPath: 'D:\python\python.exe'   # 三个路径均可留空=自动探测
    pandocPath: ''
    sofficePath: ''
    renderTimeoutSec: 90                  # 单公式 LibreOffice 渲染超时
```

## 使用示例（模型侧调用）

```jsonc
// 生成带编号的公式文档
{ "tool": "latex2office_generate",
  "args": {
    "output": "D:/docs/equations.docx",
    "title": "运动方程",
    "formulas": [
      { "latex": "\\frac{\\partial u}{\\partial t} + u\\frac{\\partial u}{\\partial x} = -\\frac{1}{\\rho}\\frac{\\partial p}{\\partial x}", "number": "1.1" }
    ]
  } }

// 插入现有文档，定位到某段文字之后
{ "tool": "latex2office_insert",
  "args": {
    "file": "D:/docs/report.docx",
    "formulas": [
      { "latex": "E = mc^2", "position": { "mode": "after_text", "text": "质能方程：" }, "number": "2.1" },
      { "latex": "\\sigma_x = \\frac{F}{A}", "position": { "mode": "end" } }
    ]
  } }
```

## 卸载 / 回滚

删除 profile manifest 中上面加的两行（dependencies + bundles），重启即完全移除。
本插件**零 npm 依赖**（纯 ES module + 随包 Python 引擎），不残留任何传递依赖。

## 崩溃安全设计

- `apply()` 只注册工具、零外部 IO，加载阶段不可能抛异常打断 harness 启动
- 所有重活（pandoc / Python / LibreOffice）走 DSH `ctx.subprocess` 服务：超时上限、stdout 字节上限 + spill 溢出保护、AbortSignal 协作中止；harness 主进程零阻塞
- 工具执行全程 try/catch，错误以 lossless JSON（`{ok:false, error:...}`）返回而非 reject
- LibreOffice 渲染使用隔离的临时用户 profile，绝不锁用户正在编辑的文档

## 常见问题

- **插入的公式在 Word/WPS 里能编辑吗？** docx 可以（原生 OMML）；pptx 默认是高清图片（兼容性优先），传 `math_mode:"omml"` 可强制原生
- **批量里有一条写错了怎么办？** 什么都不用做——原子性保证文件未被修改，修好那条重发即可
- **中文路径 / 中文锚点支持吗？** 支持，测试集覆盖
- **渲染出的图片有黑边？** 不会；白底自动转透明 + 紧裁剪 + 12px 余量

## 开发

```bash
node tests/test_load.mjs
```

隔离加载测试：模拟 cordis 上下文，验证 import → apply（3 工具注册零异常）→ generate/preview/insert 真子进程端到端 → 失败路径返回干净 JSON。产物写入系统临时目录，不触碰真实 DSH profile。依赖发现可用 `L2O_TEST_PYTHON` / `L2O_TEST_PANDOC` / `L2O_TEST_SOFFICE` 环境变量指定绝对路径。

引擎协议：`engine/engine.py` 从 stdin 读一个 JSON 请求、向 stdout 写一个 JSON 结果（成功与已处理错误均退出码 0）。

## License

MIT
