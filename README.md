# MathType 公式编号 · 章号/节号自动跟随标题

一个 Word 全局模板（`.dotm`），让 MathType 公式编号里的**章号**和**节号**自动等于所在章/节标题的编号，保存时自动同步。

支持两种编号格式：

- `(#C1.#E1)` → `(4.1)`：只有章号
- `(#C1.#S1.#E1)` → `(4.3.1)`：章号 + 节号（需在 MathType 里勾选「节编号」）

---

## 功能

| 功能 | 说明 |
|---|---|
| **章号 / 节号自动跟随** | 增、删、移动章节后，保存一次即全部重排；节号按章内序号自动重排、换章归 1 |
| **完全保留 MathType 设置** | 编号样式 `1,2,3` / `I,II,III` / `a,b,c`、括号、分隔符、附件，全部仍由 MathType「格式化公式编号」决定 |
| **交叉引用不受影响** | MathType 的「插入公式引用」、更新编号、书签跳转照常工作 |
| **自动 + 手动** | 保存 / 打印 / 另存为前自动同步；也可以随时点「立即同步编号」 |
| **可一键关闭** | 「关闭自动同步」后，保存和打印完全跳过，零耗时 |
| **全局生效** | 装在 Word 的 STARTUP 目录，所有文档自动可用，不用每篇重新配置 |
| **快** | 结构有变化约 0.5 秒，结构没变化约 0.1 秒（319 段实测） |

---

## 环境要求

- Windows 7 / 8 / 10 / 11（在 win10 上验证）
- Microsoft Word 2010–2024 或 Microsoft 365（在 365 x64 上验证）
- MathType 7.x（在 7.8 上验证）
- WPS 未测试

---

## 安装

1. 正常安装 MathType；
2. **关闭所有 Word 窗口**；
3. 右键 `install.ps1` → **使用 PowerShell 运行**；或者：

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```

装完打开 Word，在 **引用** 选项卡最右边会出现 **「公式编号」** 组。

---

## 使用

### 三个按钮

位置：**引用 → 公式编号**

| 按钮 | 作用 |
|---|---|
| **立即同步编号** | 马上同步一次；结果显示在左下角状态栏 |
| **开启自动同步** | 保存 / 打印 / 另存为前自动同步（默认状态） |
| **关闭自动同步** | 保存 / 打印时不再同步，适合文档很大的时候临时提速 |

### 日常用法

1. 正常写论文，正常用 MathType 插公式、插编号、插引用；
2. 按 `Ctrl+S` —— 编号会自动修正为当前章号与节号；
3. 增删章节后同样只需保存一次，后续所有公式的章号、节号自动重排。

同步结果只在状态栏显示，例如：

```
公式编号已同步：4 章（有更新） 20:31:07
```

---

## 卸载

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

卸载后，已写入文档的 `SEQ MTChap \r N \h` / `SEQ MTSec \r N \h` 锚点仍留在文档里（不可见、无害），编号保持最后一次同步的结果；需要清除可 `Alt+F9` 后手工删除。

> 节号是否生效取决于 MathType 的「格式化公式编号」里是否勾选了**节编号**：勾了才会出现 `SEQ MTSec` 域，插件也才会去维护节锚点；只勾章号时不会插入任何节锚点。

---

## 目录结构

```
mathtype-chapsync/
├─ MathTypeChapSync.dotm   # 已构建好的全局模板（随仓库发布）
├─ install.ps1             # 安装
├─ uninstall.ps1           # 卸载
├─ build.ps1               # 重新构建模板
└─ src/
   ├─ MTChapSync.bas       # VBA：同步引擎、命令宏、功能区回调
   ├─ ThisDocument.txt     # 模板文档模块
   └─ customUI.xml         # 功能区：引用 → 公式编号
```

## 自行构建

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

构建会复制一份 `Normal.dotm` 作为种子、写入 VBA 代码、注入功能区 XML，输出到 `dist\MathTypeChapSync.dotm`。

> 注意：构建需要临时打开 VBA 工程访问权限（`AccessVBOM`），脚本会在结束后恢复原值。

---