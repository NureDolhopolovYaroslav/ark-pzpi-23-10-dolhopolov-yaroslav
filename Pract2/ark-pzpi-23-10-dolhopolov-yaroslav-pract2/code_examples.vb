' Приклад для методу Replace Temp With Query
' Код до рефакторингу
Public Class OrderCalculator
    ' Проблема: тимчасова змінна basePrice захаращує логіку
    Public Function CalculateTotalOld(price As Decimal, quantity As Integer) As Decimal
        Dim basePrice As Decimal = price * quantity ' Зайва тимчасова змінна
        If basePrice > 1000 Then
            Return basePrice * 0.95
        Else
            Return basePrice * 0.98
        End If
    End Function
End Class

'Код після рефакторингу
Public Class OrderCalculatorRefactored
    ' Рішення: логіка обчислення винесена в окремий метод-запит
    Public Function CalculateTotalNew(price As Decimal, quantity As Integer) As Decimal
        If BasePrice(price, quantity) > 1000 Then
            Return BasePrice(price, quantity) * 0.95
        Else
            Return BasePrice(price, quantity) * 0.98
        End If
    End Function

    ' Новий метод-запит, що інкапсулює логіку розрахунку
    Private Function BasePrice(price As Decimal, quantity As Integer) As Decimal
        Return price * quantity
    End Function
End Class

' ==================================================================================

' Приклад для методу Replace Nested Conditional with Guard Clauses
' Код до рефакторингу
Public Class DiscountService
    ' Проблема: "піраміда смерті" - глибоко вкладені умови
    Public Function GetDiscountPercentageOld(customer As Customer) As Decimal
        If customer IsNot Nothing Then
            If customer.IsActive Then
                If customer.TotalPurchases > 10000 Then
                    Return 0.1 ' 10%
                Else
                    Return 0.05 ' 5%
                End If
            Else
                Return 0 ' Неактивний клієнт
            End If
        Else
            Throw New ArgumentNullException(NameOf(customer))
        End If
    End Function
End Class

Public Class Customer
    Public Property IsActive As Boolean
    Public Property TotalPurchases As Decimal
End Class

' Код після рефакторингу 
Public Class DiscountServiceRefactored
    ' Рішення: захисні клаузи на початку роблять логіку лінійною
    Public Function GetDiscountPercentageNew(customer As Customer) As Decimal
        ' Захисна клауза 1: перевірка на null
        If customer Is Nothing Then
            Throw New ArgumentNullException(NameOf(customer))
        End If

        ' Захисна клауза 2: перевірка на активність
        If Not customer.IsActive Then
            Return 0
        End If

        ' Основний шлях виконання - тепер він чистий і зрозумілий
        If customer.TotalPurchases > 10000 Then
            Return 0.1
        End If

        Return 0.05
    End Function
End Class

' ==================================================================================

' Приклад для методу Preserve Whole Object
' Код до рефакторингу
Public Class CustomerService
    ' Проблема: довгий список параметрів, що належать одній сутності
    Public Sub UpdateCustomerProfileOld(customerId As Integer,
                                        street As String,
                                        city As String,
                                        postalCode As String,
                                        phone As String)
        ' Логіка оновлення профілю
        Console.WriteLine($"Оновлення адреси: {street}, {city}, {postalCode}")
    End Sub
End Class

' Код після рефакторингу

' Створюємо клас для групування пов'язаних даних
Public Class CustomerAddress
    Public Property Street As String
    Public Property City As String
    Public Property PostalCode As String

    ' Можна додати корисні методи
    Public Overrides Function ToString() As String
        Return $"{Street}, {City}, {PostalCode}"
    End Function
End Class

Public Class CustomerServiceRefactored
    ' Рішення: передаємо цілісний об'єкт замість набору параметрів
    Public Sub UpdateCustomerProfileNew(customerId As Integer,
                                        address As CustomerAddress,
                                        phone As String)
        ' Логіка стала набагато чистішою
        Console.WriteLine($"Оновлення адреси: {address.ToString()}")
    End Sub

    ' Приклад виклику методу
    Public Sub DemoMethodCall()
        ' Створення об'єкта адреси
        Dim customerAddress As New CustomerAddress() With {
            .Street = "вул. Наукова, 42",
            .City = "Харків",
            .PostalCode = "61152"
        }

        ' Виклик методу з одним об'єктом замість трьох параметрів
        UpdateCustomerProfileNew(123, customerAddress, "+380999999999")
    End Sub
End Class