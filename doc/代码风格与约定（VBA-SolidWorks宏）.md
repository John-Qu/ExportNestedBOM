目标
- 统一 VBA/SolidWorks 宏的编码风格，提升可读性与维护性。

命名约定
- 模块：功能名或域名（如 `UpdateProperties`、`Utils`）。
- 过程/函数：动词开头（`Run_AddCustomProps`、`BuildPropArrays`）。
- 变量：驼峰或下划线，语义清晰（避免单字母除循环变量）。
- 常量：全大写加下划线（`swDocPART`、`swCustomInfoText`）。

错误处理
- 外部 API 调用前后使用 `On Error Resume Next` → 立即 `On Error GoTo 0` 收敛；仅包裹最小范围。
- 用户操作路径（打开/保存）尽量提示而非静默失败；必要时弹窗说明。

写入策略
- 自定义属性与配置特定属性采用“先删后增”（`Add2`，文本类型），保证一致性与幂等性。
- 读取采用“两段式”：先配置特定，再回退文档级（`Get6`）。

UI 约定
- `frmProps` 行项目动态生成：名称 | 值（可编辑） | 勾选导入。
- 默认勾选导入；支持“全选/全不选”。

工具函数
- 文本安全化：`SafeStr`；
- 路径解析：`BaseNameNoExt`、`ParentFolderName`；
- 解析规则：`ParseCodeAndName`、`ParseProjectFromFolder`；
- 质量与识别：`GetMassKgString`、`IsSheetMetal`。

代码组织
- 将“预清理”（删除命中关键词方程式/属性）与“写入”分离；主流程清晰：选择范围 → 构建数组 → 弹窗 → 写入 → 保存/统计。

注释与文档
- 顶部模块注释给出用途与入口；重要过程前说明输入/输出与副作用。
- 变更在 `doc/变更日志.md` 记录，发布流程按 `doc/版本发布流程.md`。