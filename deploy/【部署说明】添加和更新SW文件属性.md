可以的。下面在“极简版”基础上为你加上新需求，并把“确认窗口”升级为可就地编辑、逐项勾选（默认全选）的对话框。

新增与改动要点
- 新增属性：设计、定型日期、型号、SUPPLIER
- 取值规则：
  - 设计：优先读取文件中的“设计”属性值；若不存在，则取自操作系统用户名（Windows 登录名）
  - 定型日期：取自Solidworks内部的文件创建日期（yyyy-mm-dd）
  - 型号、SUPPLIER：优先从文档现有的“自定义属性”读取；若不存在，则在确认界面中显示为空，默认勾选，方便填写后写入
- 确认界面升级：
  - 每个属性一行：属性名（不可改） + 值（可编辑） + “导入”勾选（默认选中）
  - 支持就地编辑：修改值后点击“添加/更新”即写入
  - 逐文件操作按钮：添加/更新、跳过、取消（取消终止批量）
- 仍保留原逻辑：文件名/文件夹名解析、材料/质量读取、钣金识别、批量处理、已存在属性“更新”而非重复添加、处理完成自动保存

使用前准备
- 在 SolidWorks 打开“宏（VBA）编辑器”，建议创建一个宏工程（.swp）
- 添加一个标准模块 Module1，把下面“模块代码”完整粘贴进去
- 新建一个 UserForm，命名为 frmProps，并按下列控件布局（非常简单）：
  - 在 frmProps 上放置：
    - Label：Name = lblInfo，Caption 随意（用于显示文件信息）
    - Frame：Name = fraList，SpecialEffect = fmSpecialEffectEtched，ScrollBars = fmScrollBarsVertical，KeepScrollBarsVisible = fmScrollBarsVertical，Caption 留空
    - CommandButton：Name = cmdOK，Caption = 添加/更新
    - CommandButton：Name = cmdSkip，Caption = 跳过
    - CommandButton：Name = cmdCancel，Caption = 取消
    - CheckBox（可选）：Name = chkAll，Caption = 全选/全不选
  - 大致布局建议（像素/点数不苛刻）：  
    - frmProps.Width ≈ 560，Height ≈ 600  
    - lblInfo.Top = 10，Left = 10，Width ≈ 520  
    - fraList.Top = 50，Left = 10，Width ≈ 520，Height ≈ 460  
    - chkAll.Top = 515，Left = 10  
    - cmdOK.Top = 545，Left = 260  
    - cmdSkip.Top = 545，Left = 360  
    - cmdCancel.Top = 545，Left = 460  
  - 不需要在设计期添加任何行项目，行项目（属性名/文本框/勾选）会由代码动态生成

模块代码（Module1），从复制粘贴。
- 从 SW_src/UpdateProperties.bas 完整复制粘贴。

UserForm 代码（frmProps）
- 从 SW_src/UpdateProperties_UserForm.bas 完整复制粘贴。

使用说明
- 运行宏入口：Run_AddCustomProps
- 选择“当前文件”或“文件夹批量”
- 在弹出的“属性确认”窗口里：
  - 直接编辑“值”列的内容
  - 取消勾选某一项则本轮对该属性不写入（不改动已有值，也不新增）
  - 点击“添加/更新”：按勾选写入，并对批量处理自动保存文件
  - 点击“跳过”：不写入该文件
  - 点击“取消”：终止剩余批量

如果你希望我把 UserForm 的控件也改成纯代码动态创建（无需设计器放控件），或要我开启子文件夹递归、批量不逐个弹窗等，我可以继续帮你调整。