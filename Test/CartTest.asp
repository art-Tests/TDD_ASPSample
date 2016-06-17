<!-- #include virtual="/Model/Cart.asp" -->
<%
Class CartTest ' Extends TestCase

  Private target

  Public Function TestCaseNames()
    TestCaseNames = Array("Buy_Book1_For_1_Should_Retun_100","Buy_Book1_For_1_And_Book2_For2_Should_Retun_190")
  End Function
  
  Public Sub SetUp()
    
  End Sub
  
  Public Sub TearDown()
    Set target = Nothing
  End Sub
  

  Public Sub Buy_Book1_For_1_Should_Retun_100(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("哈利波特1", 100, 1)
    dim Books
    Books = Array(book1)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 100
    oTestResult.AssertEquals expected, actual, "價格不同！"
  End Sub

  Public Sub Buy_Book1_For_1_And_Book2_For2_Should_Retun_190(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("哈利波特2", 100, 2)
    dim book2 : Set book2 = (New Book)("哈利波特1", 100, 1)
    dim Books
    Books = Array(book1,book2)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 190
    oTestResult.AssertEquals expected, actual, "價格不同！"
  End Sub

End Class
%>