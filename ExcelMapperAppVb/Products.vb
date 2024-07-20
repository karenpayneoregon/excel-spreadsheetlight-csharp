Public Class Products
    Public Property ProductID As Integer

    Public Property ProductName As String

    Public Property CategoryName As String
    Public Property SupplierID As Integer?

    Public Property CategoryID As Integer?

    Public Property Supplier As String
    Public Property QuantityPerUnit As String

    Public Property UnitPrice As Decimal?

    Public Property UnitsInStock As Short?

    Public Property UnitsOnOrder As Short?

    Public Property ReorderLevel As Short?

    Public Overrides Function ToString() As String
        Return ProductName
    End Function

End Class
