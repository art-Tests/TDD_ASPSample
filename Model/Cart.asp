<!-- #include virtual="/Model/Book.asp" -->
<%
'購物車'
Class Cart
    '建構式'
    public default function Init()
        Set Init = Me
    end function

    public function Pay(iBooks)
        Pay = 0
    end function
End Class
%>