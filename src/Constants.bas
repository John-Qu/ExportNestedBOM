' Constants.bas - System constants and configuration
Option Explicit

' 系统配置常量
Public Const MAX_RECURSION_DEPTH As Integer = 10
Public Const MAX_LOG_SIZE_MB As Long = 5
Public Const TEMP_FILE_PREFIX As String = "ExportNestedBOM_"

' SolidWorks文档类型常量
Public Const swDocPART As Integer = 1
Public Const swDocASSEMBLY As Integer = 2
Public Const swDocDRAWING As Integer = 3

' 文件扩展名
Public Const DRAWING_EXT As String = ".slddrw"
Public Const ASSEMBLY_EXT As String = ".sldasm"
Public Const PART_EXT As String = ".sldprt"

' 默认列名映射
Public Function GetQuantityColumnNames() As Variant
    GetQuantityColumnNames = Array("数量", "QTY", "Qty", "QUANTITY", "数量(QTY)")
End Function

Public Function GetNameColumnNames() As Variant
    GetNameColumnNames = Array("名称", "PART NAME", "Name", "零件名称", "品名")
End Function

Public Function GetPartNumberColumnNames() As Variant
    GetPartNumberColumnNames = Array("代号", "PART NUMBER", "Part Number", "PARTPATH", "零件路径", "零件号", "图号")
End Function

Public Function GetAssemblyColumnNames() As Variant
    GetAssemblyColumnNames = Array("是否组装", "Is Assembly", "组装", "是否组件", "IS ASSEMBLY", "组装体", "ASSEMBLY")
End Function

' 是否组装判断值
Public Function GetAssemblyTrueValues() As Variant
    GetAssemblyTrueValues = Array("是", "Y", "YES", "TRUE", "1", "组装", "装配")
End Function

' 错误码定义
Public Const ERR_SOLIDWORKS_NOT_FOUND As Long = 1001
Public Const ERR_FILE_NOT_FOUND As Long = 1002
Public Const ERR_NO_BOM_TABLE As Long = 1003
Public Const ERR_CIRCULAR_REFERENCE As Long = 1004
Public Const ERR_DIRECTORY_NOT_WRITABLE As Long = 1005
Public Const ERR_RECURSION_DEPTH_EXCEEDED As Long = 1006