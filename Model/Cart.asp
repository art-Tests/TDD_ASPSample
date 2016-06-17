<!-- #include virtual="/Model/Book.asp" -->
<%
'購物車'
Class Cart
    '建構式'
    public default function Init()
        Set Init = Me
    end function

    public function Pay(iBooks)
        '建立欄位，資料型態請參考http://www.w3schools.com/asp/ado_datatypes.asp
        dim bookTypeCnt,discount,result,thisGroupPay,isRunNext
        dim rs : set rs = server.createobject("adodb.recordset")
        rs.Fields.Append "Name" , 202 , 10
        rs.Fields.Append "Price" , 5
        rs.Fields.Append "No" , 2 
        rs.Open

        dim bookGroup
        isRunNext = false

        do
            set bookGroup = server.createobject("adodb.recordset")
            bookGroup.Fields.Append "Name" , 202 , 10
            bookGroup.Fields.Append "Price" , 5
            bookGroup.Fields.Append "No" , 2 
            bookGroup.Open
            isRunNext = getBookRs(isRunNext,iBooks,rs,bookGroup)
            result = result + getBookGroupPrice(bookGroup)
            'call showRs(rs)
            set bookGroup = nothing
        loop until not isRunNext

        pay = result


    end function


    private function getBookGroupPrice(ByRef bookGroup)
        dim bookTypeCnt,discount
        'response.write "<hr/>"
        'call showRs(bookGroup)
        bookTypeCnt = getBookTypeCntByAllBooks(bookGroup)  '取得書籍種類數量'
        discount    = getDiscountByTypeCnt(bookTypeCnt) '取得本次購物的折扣'
        'response.write "BookGroupPrice:<span style='color:red'>"&getTotalPrice(bookGroup,discount)&"</span><hr/>"
        getBookGroupPrice = getTotalPrice(bookGroup,discount)    '每一本書的價格合計'
    end function

    private sub showRs(ByRef bookGroup)
        'response.write "RS CNT:"&bookGroup.recordcount&"<br/>"
        bookGroup.moveFirst
        if not bookGroup.EOF then
            while not bookGroup.EOF
                'response.write "Name:["&bookGroup("Name")&"] No:["&bookGroup("No")&"] Price:["&bookGroup("Price")&"]<br/>"
                bookGroup.movenext
            wend
        end if
        bookGroup.moveFirst
    end sub


    private function getSinglePriceBy(originPrice,discount)
        getSinglePriceBy = originPrice * discount
    end function

    private function getTotalPrice(ByRef bookGroup,discount)
        dim result : result = 0
        dim i,book,thisBookPrice
        bookGroup.moveFirst

        while not bookGroup.EOF
            thisBookPrice = getSinglePriceBy(bookGroup("Price"),discount)
            result = thisBookPrice + result
            bookGroup.movenext
        wend
        getTotalPrice = result
    end function

    private function getDiscountByTypeCnt(bookTypeCnt)
        dim discount
        select case bookTypeCnt
            case 2      : discount = 0.95
            case 3      : discount = 0.9
            case 4      : discount = 0.8
            case 5      : discount = 0.75
            case else   : discount = 1
        end select
        getDiscountByTypeCnt = discount
    end function

    private function getBookTypeCntByAllBooks(ByRef rs)
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
        getBookTypeCntByAllBooks = bookTypeCnt
    end function

    private function getBookRs(isFirst,iBooks, ByRef rs,ByRef bookGroup)

        if not isFirst then
            dim i
            for i = lbound(iBooks) to ubound(iBooks)
                rs.AddNew
                rs("Name") = iBooks(i).GetName
                rs("Price") = iBooks(i).GetPrice
                rs("No") = iBooks(i).GetNo
                rs.Update
                rs.Movenext
                'response.write "Name:["&rs("Name")&"] No:["&rs("No")&"] Price:["&rs("Price")&"]<br/>"
            next 
        end if

        rs.Sort = "No"
        rs.moveFirst
        dim bookTypeCnt :   bookTypeCnt = 0
        dim nowType :   nowType = 0
        dim lastType :  lastType = 0
        while not rs.EOF
            nowType = rs("No")
            if(lastType<>nowType) then 
                bookGroup.AddNew
                bookGroup("Name") = rs("Name")
                bookGroup("Price") = rs("Price")
                bookGroup("No") = rs("No")
                bookGroup.Update
                bookGroup.Movenext
                rs.delete
                lastType=nowType
            end if
            'response.write "Name:["&rs("Name")&"] No:["&rs("No")&"] Price:["&rs("Price")&"]<br/>"
            rs.movenext
        wend
        if rs.recordcount>0 then
            getBookRs = true
        else
            getBookRs = false
        end if

    end function
End Class


%>