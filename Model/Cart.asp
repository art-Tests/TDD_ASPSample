<!-- #include virtual="/Model/Book.asp" -->
<%
'購物車'
Class Cart
    '建構式'
    public default function Init()
        Set Init = Me
    end function

    public function Pay(iBooks)
        dim bookTypeCnt,discount
        bookTypeCnt = getBookTypeCntByAllBooks(iBooks)  '取得書籍種類數量'
        discount    = getDiscountByTypeCnt(bookTypeCnt) '取得本次購物的折扣'
        Pay         = getTotalPrice(iBooks,discount)    '每一本書的價格合計'
    end function

    private function getSinglePriceBy(originPrice,discount)
        getSinglePriceBy = originPrice * discount
    end function

    private function getTotalPrice(ibooks,discount)
        dim result : result = 0
        dim i,book,thisBookPrice
        for i = lbound(iBooks) to ubound(iBooks)
            thisBookPrice = getSinglePriceBy(iBooks(i).GetPrice,discount)
            result = thisBookPrice + result
        next 
        getTotalPrice = result
    end function

    private function getDiscountByTypeCnt(bookTypeCnt)
        dim discount
        select case bookTypeCnt
            case 2      : discount = 0.95
            case 3      : discount = 0.9
            case 4      : discount = 0.8
            case else   : discount = 1
        end select
        getDiscountByTypeCnt = discount
    end function

    private function getBookTypeCntByAllBooks(iBooks)
        dim rs : set rs = server.createobject("adodb.recordset")
        '建立欄位，資料型態請參考http://www.w3schools.com/asp/ado_datatypes.asp
        rs.Fields.Append "Name" , 202 , 10
        rs.Fields.Append "Price" , 5
        rs.Fields.Append "No" , 2 
        rs.Open
        dim i
        for i = lbound(iBooks) to ubound(iBooks)
            rs.AddNew
            rs("Name") = iBooks(i).GetName
            rs("Price") = iBooks(i).GetPrice
            rs("No") = iBooks(i).GetNo
            rs.Update
            rs.Movenext
        next 
        rs.Sort = "No"
        rs.moveFirst
        dim bookTypeCnt :   bookTypeCnt = 0
        dim nowType :   nowType = 0
        dim lastType :  lastType = 0
        while not rs.EOF
            nowType = rs("No")
            if(lastType<>nowType) then 
                bookTypeCnt = bookTypeCnt + 1
                lastType = nowType
            end if
            'response.write "Name:["&rs("Name")&"] No:["&rs("No")&"] Price:["&rs("Price")&"]<br/>"
            rs.movenext
        wend
        'response.write "BookTypeCnt:["&bookTypeCnt&"]<br/>"
        rs.close
        set rs = nothing
        getBookTypeCntByAllBooks = bookTypeCnt
    end function

End Class


%>