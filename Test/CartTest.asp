<!-- #include virtual="/Model/Cart.asp" -->
<%
Class CartTest ' Extends TestCase

  Private target

  Public Function TestCaseNames()
    TestCaseNames = Array("Buy_Book1_For_1_Should_Retun_100","Buy_Book1_For_1_And_Book2_For2_Should_Retun_190","Buy_Book123_For_1_Should_Retun_270","Buy_Book1234_For_1_Should_Retun_320","Buy_Book12345_For_1_Should_Retun_375","Buy_Book12_For_1_And_Book3_for_2_Should_Retun_370")
  End Function
  
  Public Sub SetUp()
    
  End Sub
  
  Public Sub TearDown()
    Set target = Nothing
  End Sub
  

  Public Sub Buy_Book1_For_1_Should_Retun_100(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭1", 100, 1)
    dim Books
    Books = Array(book1)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 100
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub

  Public Sub Buy_Book1_For_1_And_Book2_For2_Should_Retun_190(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭2", 100, 2)
    dim book2 : Set book2 = (New Book)("猧疭1", 100, 1)
    dim Books
    Books = Array(book1,book2)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 190
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub

  Public Sub Buy_Book123_For_1_Should_Retun_270(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭2", 100, 2)
    dim book2 : Set book2 = (New Book)("猧疭1", 100, 1)
    dim book3 : Set book3 = (New Book)("猧疭3", 100, 3)
    dim Books
    Books = Array(book1,book2,book3)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 270
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub

  Public Sub Buy_Book1234_For_1_Should_Retun_320(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭2", 100, 2)
    dim book2 : Set book2 = (New Book)("猧疭1", 100, 1)
    dim book3 : Set book3 = (New Book)("猧疭3", 100, 3)
    dim book4 : Set book4 = (New Book)("猧疭4", 100, 4)
    dim Books
    Books = Array(book1,book2,book3,book4)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 320
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub

  Public Sub Buy_Book12345_For_1_Should_Retun_375(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭2", 100, 2)
    dim book2 : Set book2 = (New Book)("猧疭1", 100, 1)
    dim book3 : Set book3 = (New Book)("猧疭3", 100, 3)
    dim book4 : Set book4 = (New Book)("猧疭4", 100, 4)
    dim book5 : Set book5 = (New Book)("猧疭5", 100, 5)
    dim Books
    Books = Array(book1,book2,book3,book4,book5)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 375
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub

  Public Sub Buy_Book12_For_1_And_Book3_for_2_Should_Retun_370(oTestResult)
    'arrange'
    dim book1 : Set book1 = (New Book)("猧疭1", 100, 1)
    dim book2 : Set book2 = (New Book)("猧疭2", 100, 2)
    dim book3 : Set book3 = (New Book)("猧疭3", 100, 3)
    dim book4 : Set book4 = (New Book)("猧疭3", 100, 3)
    dim Books
    Books = Array(book1,book2,book3,book4)

    set target = new Cart
    'act'
    dim actual : actual = target.Pay(Books)
    'assert'
    dim expected : expected = 370
    oTestResult.AssertEquals expected, actual, "基ぃ"
  End Sub
End Class
%>