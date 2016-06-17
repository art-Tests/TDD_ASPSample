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
        bookTypeCnt = 0
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

        select case bookTypeCnt
            case 2      : discount = 0.95
            case else   : discount = 1
        end select


        dim result : result = 0
        dim book,thisBookPrice
        for i = lbound(iBooks) to ubound(iBooks)
            thisBookPrice = (iBooks(i).GetPrice*discount)
            result = thisBookPrice + result
        next 
        pay = result
    end function


End Class


%>