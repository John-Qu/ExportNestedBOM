目的
- 提供最小可操作的检查清单，确保在无外网环境下也能手工部署并运行 `Run_AddCustomProps`。

准备
- SolidWorks 2019 已安装并可打开宏编辑器（VBA）。
- 目标模型文件存在（`*.sldprt`/`*.sldasm`）。

步骤
- 打开 SolidWorks → 工具 → 宏 → 新建（建议创建 `.swp` 宏工程）。
- 添加标准模块 `Module1`：将 `SW_src/UpdateProperties.bas` 的内容完整粘贴。
- 新建 `UserForm` 命名为 `frmProps`，添加控件：
  - `Label`：Name = `lblInfo`，Caption 随意（用于显示文件信息）
  - `Frame`：`fraList`（滚动条启用），SpecialEffect = fmSpecialEffectEtched，ScrollBars = fmScrollBarsVertical，KeepScrollBarsVisible = fmScrollBarsVertical，Caption 留空
  - CommandButton：Name = `cmdOK`，Caption = 添加/更新
  - CommandButton：Name = `cmdSkip`，Caption = 跳过
  - CommandButton：Name = `cmdCancel`，Caption = 取消
  - CheckBox（可选）：Name = `chkAll`，Caption = 全选/全不选
  - 大致布局建议（像素/点数不苛刻）：  
    - frmProps.Width ≈ 560，Height ≈ 600  
    - lblInfo.Top = 10，Left = 10，Width ≈ 520  
    - fraList.Top = 50，Left = 10，Width ≈ 520，Height ≈ 460  
    - chkAll.Top = 515，Left = 10  
    - cmdOK.Top = 545，Left = 260  
    - cmdSkip.Top = 545，Left = 360  
    - cmdCancel.Top = 545，Left = 460  
  - 不需要在设计期添加任何行项目，行项目（属性名/文本框/勾选）会由代码动态生成
- 在 `frmProps` 代码窗口粘贴 `SW_src/UpdateProperties_UserForm.bas` 的内容。
- 保存宏工程。

快速自检
- 打开一个 Part 或 Assembly。
- 运行宏入口 `Run_AddCustomProps`：
  - 选择“当前文件”，弹出确认窗口；行项目按代码动态生成。
  - 点击“添加/更新”，完成写入并提示。
- 打开自定义属性与配置特定属性窗口，核对写入结果。

常见误区
- `frmProps` 名称不一致 → 需确保窗体名与代码一致。
- 未粘贴 `UserForm` 代码 → 动态行创建不会发生。
- 文件只读/权限不足 → 写入失败；需要解除只读或以可写路径保存。
- 在 SolidWorks 之外运行宏 → 无法获取 `Application.SldWorks`。

回滚与卸载
- 删除宏工程或移除新增模块与窗体；不会影响模型文件本身。