<%
Class Book

    Private Name
    Private Price
    Private No

    public default function Init(iName, iPrice, iNo)
        Name = iName
        Price = iPrice
        No = iNo
        Set Init = Me
    end function

    Public Function GetName()
        GetName = Name
    End Function

    Public Function GetPrice()
        GetPrice = Price
    End Function

    Public Function GetNo()
        GetNo = No
    End Function

End Class
%>