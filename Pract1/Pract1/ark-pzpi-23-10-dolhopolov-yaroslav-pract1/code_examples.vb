' ======= В.1 =======
' Стандартна структура проекту Visual Basic

' MySolution.sln (файл рішення)
' MyProject.vbproj (файл проєкту)

Imports System.Net.Mime.MediaTypeNames

Module Program
    Sub Main()
        Application.Run(New MainForm())
    End Sub
End Module

Public Class MainForm
    Inherits Form

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Ініціалізація компонентів форми
    End Sub
End Class


' ======= В.2 =======
' Організація коду через Namespace та Region

Namespace MyCompany.MyApplication.BusinessLogic

    Public Class CustomerService
        Private _customers As List(Of Customer)

        Public Sub New()
            _customers = New List(Of Customer)()
        End Sub

#Region "Public Methods"
        Public Sub AddCustomer(customer As Customer)
            ValidateCustomer(customer)
            _customers.Add(customer)
        End Sub

        Public Function GetCustomerById(id As Integer) As Customer
            Return _customers.FirstOrDefault(Function(c) c.Id = id)
        End Function
#End Region

#Region "Private Methods"
        Private Sub ValidateCustomer(customer As Customer)
            If customer Is Nothing Then
                Throw New ArgumentNullException(NameOf(customer))
            End If
        End Sub
#End Region
    End Class

End Namespace


' ======= В.3 =======
' Форматування коду: погано та добре

' Погано:
Public Class BadFormatting
    Public Function Calculate(a As Decimal, b As Integer) As Decimal
        Return a * b
    End Function
End Class

' Добре:
Public Class GoodFormatting
    Public Function Calculate(
        price As Decimal,
        quantity As Integer) As Decimal

        Return price * quantity
    End Function

    Public Sub ProcessUserData(
        userName As String,
        userAge As Integer,
        userEmail As String)

        If Not String.IsNullOrEmpty(userName) AndAlso
           userAge > 0 AndAlso
           Not String.IsNullOrEmpty(userEmail) Then

            ' Обробка даних
        End If
    End Sub
End Class


' ======= В.4 =======
' Конвенції іменування

' Добре (відповідно до конвенцій Visual Basic)
Public Class CustomerRepository    ' PascalCase для класів
    Private _customerList As List(Of Customer)    ' camelCase з префіксом _
    Private Const MaxOrderCount As Integer = 100  ' PascalCase для констант

    Public Function GetActiveCustomers() As List(Of Customer)
        Return _customerList.Where(
            Function(c) c.IsActive AndAlso
                       c.OrderCount < MaxOrderCount).ToList()
    End Function
End Class

' Погано (порушення конвенцій)
Public Class customer_repository    ' snake_case — не прийнято
    Private CustomerList As List(Of Customer)    ' PascalCase для поля (погано)
    Private Const max_order_count As Integer = 100    ' camelCase для константи (погано)

    Public Function get_active_customers() As List(Of Customer)    ' snake_case
        Return CustomerList.Where(Function(c) c.is_active).ToList()
    End Function
End Class


' ======= В.5 =======
' Уникнення "магічних чисел"

' Добре: використання констант
Public Class OrderValidator
    Private Const MinimumAge As Integer = 18
    Private Const MaxOrderAmount As Decimal = 10000D

    Public Function ValidateCustomerOrder(customer As Customer, orderAmount As Decimal) As Boolean
        If customer.Age >= MinimumAge AndAlso orderAmount <= MaxOrderAmount Then
            Return True
        End If
        Return False
    End Function
End Class

' Погано: магічні числа прямо в коді
Public Class BadOrderValidator
    Public Function ValidateCustomerOrder(customer As Customer, orderAmount As Decimal) As Boolean
        If customer.Age >= 18 AndAlso orderAmount <= 10000.0 Then
            Return True
        End If
        Return False
    End Function
End Class


' ======= В.6 =======
' Коментарі: пояснюємо "чому", а не "що"

' Погано: пояснюємо "що"
' Додаємо 10 до ціни
price += 10

' Добре: пояснюємо "чому"
' Застосовуємо стандартну надбавку згідно з ціновою політикою компанії
price += 10

' Позначення майбутніх поліпшень:
' TODO: додати валідацію email адреси
' FIXME: виправити обробку null значень
' OPTIMIZE: покращити продуктивність


' ======= В.7 =======
' Приклад юніт-тестів (MSTest)

<TestClass>
Public Class CalculatorTests

    <TestMethod>
    Public Sub Add_TwoPositiveNumbers_ReturnsCorrectSum()
        Dim calculator As New Calculator()
        Dim expected As Integer = 5

        Dim actual As Integer = calculator.Add(2, 3)

        Assert.AreEqual(expected, actual)
    End Sub

    <TestMethod>
    <ExpectedException(GetType(ArgumentOutOfRangeException))>
    Public Sub Divide_ByZero_ThrowsArgumentOutOfRangeException()
        Dim calculator As New Calculator()

        calculator.Divide(10, 0)
    End Sub

End Class


' ======= В.8 =======
' Обробка помилок: Try-Catch

Public Class DataProcessor
    Public Function ProcessFile(filePath As String) As String
        Try
            Using reader As New System.IO.StreamReader(filePath)
                Dim content As String = reader.ReadToEnd()
                Return ProcessContent(content)
            End Using
        Catch ex As System.IO.FileNotFoundException
            MessageBox.Show($"Файл не знайдено: {filePath}")
            Return String.Empty
        Catch ex As System.UnauthorizedAccessException
            MessageBox.Show("Немає доступу до файлу")
            Return String.Empty
        Catch ex As Exception
            MessageBox.Show($"Сталася помилка: {ex.Message}")
            Return String.Empty
        End Try
    End Function
End Class


' ======= В.9 =======
' Юніт-тестування (продовження)

<TestClass>
Public Class PrimeCheckerTests

    <TestMethod>
    Public Sub IsPrime_PrimeNumber_ReturnsTrue()
        Dim primeChecker As New PrimeChecker()

        Dim result As Boolean = primeChecker.IsPrime(17)

        Assert.IsTrue(result)
    End Sub

End Class


' ======= В.10 =======
' Використання інструментів аналізу коду (Code Analysis)

Public Class CodeAnalysisExample
    ' CA1823: Уникати невикористаних приватних полів
    Private _unusedField As String

    ' CA1051: Не оголошуйте публічні поля, краще використовуйте властивості
    Public BadField As Integer

    ' CA1062: Перевірка аргументів публічних методів на null
    Public Sub ProcessData(data As String)
        If data Is Nothing Then
            Throw New ArgumentNullException(NameOf(data))
        End If

        ' Подальша обробка data
    End Sub
End Class


' ======= В.11 =======
' Поганий код – антипатерни

Public Class BadCodeExample

    ' Неявне оголошення змінних (варто явно вказувати тип)
    Public Sub ProcessUser()
        On Error Resume Next ' Застарілий спосіб обробки помилок

        Dim x = 10        ' Неявний тип (Option Strict може заборонити)
        Dim y = "20"
        Dim z = x + y     ' Неявне перетворення типів

        ' Використання магічних чисел
        If z > 30 Then
            MessageBox.Show("Помилка")
        End If
    End Sub

    ' Занадто довгий метод (погана практика)
    Public Sub DoEverything()
        ' >100 рядків коду – важко читати і підтримувати
    End Sub

End Class


' ======= В.12 =======
' Добрий код – сучасні практики

Public Class GoodCodeExample
    Private ReadOnly _userRepository As IUserRepository
    Private Const MaxLoginAttempts As Integer = 3

    Public Sub New(userRepository As IUserRepository)
        _userRepository = userRepository
    End Sub

    Public Function AuthenticateUser(username As String, password As String) As AuthenticationResult
        Try
            If String.IsNullOrEmpty(username) OrElse String.IsNullOrEmpty(password) Then
                Return AuthenticationResult.InvalidInput
            End If

            Dim user = _userRepository.GetUserByUsername(username)
            If user Is Nothing Then
                Return AuthenticationResult.UserNotFound
            End If

            If user.LoginAttempts >= MaxLoginAttempts Then
                Return AuthenticationResult.AccountLocked
            End If

            If VerifyPassword(password, user.PasswordHash) Then
                user.LoginAttempts = 0
                _userRepository.UpdateUser(user)
                Return AuthenticationResult.Success
            Else
                user.LoginAttempts += 1
                _userRepository.UpdateUser(user)
                Return AuthenticationResult.InvalidPassword
            End If

        Catch ex As Exception
            ' Логування помилки
            Return AuthenticationResult.ServerError
        End Try
    End Function

    Private Function VerifyPassword(inputPassword As String, storedHash As String) As Boolean
        ' Тут має бути реальна логіка перевірки пароля
        Return True
    End Function
End Class

Public Enum AuthenticationResult
    Success
    InvalidInput
    UserNotFound
    InvalidPassword
    AccountLocked
    ServerError
End Enum
