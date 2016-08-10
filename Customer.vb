Public Class Customer


    'Attributes which define the customer information

    Private intCustomerID As Integer
    Private strFirstName As String
    Private strSurname As String
    Private strAddress As String
    Private strTown As String
    Private strCounty As String
    Private strPostCode As String
    Private strPhone As String
    Private strEmail As String

    'Properties to set a attribute values and return a value from an attribute

    Public Property CustomerID As Integer

        Get

            Return intCustomerID

        End Get

        Set(ByVal value As Integer)

            intCustomerID = value

        End Set

    End Property

    Public Property FirstName As String

        Get

            Return strFirstName

        End Get

        Set(ByVal value As String)

            strFirstName = value

        End Set

    End Property

    Public Property Surname As String

        Get

            Return strSurname

        End Get

        Set(ByVal value As String)

            strSurname = value

        End Set

    End Property

    Public Property Address As String

        Get

            Return strAddress

        End Get

        Set(ByVal value As String)

            strAddress = value

        End Set

    End Property

    Public Property Town As String

        Get

            Return strTown

        End Get

        Set(ByVal value As String)

            strTown = value

        End Set

    End Property

    Public Property County As String

        Get

            Return strCounty

        End Get

        Set(ByVal value As String)

            strCounty = value

        End Set

    End Property

    Public Property PostCode As String

        Get

            Return strPostCode

        End Get

        Set(ByVal value As String)

            strPostCode = value

        End Set

    End Property

    Public Property Phone As String

        Get

            Return strPhone

        End Get

        Set(ByVal value As String)

            strPhone = value

        End Set

    End Property

    Public Property Email As String

        Get

            Return strEmail

        End Get

        Set(ByVal value As String)

            strEmail = value

        End Set

    End Property

    'Default Constructor

    Sub New()

        'Sets default values for each of the properties highlighted above

        CustomerID = 1234
        FirstName = "dftFirstName"
        Surname = "dftSurname"
        Address = "dftAddress"
        Town = "dftTown"
        County = "dftCounty"
        PostCode = "dftPostCode"
        Phone = "dftPhone"
        Email = "dft@Email"

    End Sub

    'Overloaded Constructor 

    'Allows new customer values to be inserted in from external Class code

    Sub New(ByVal newCustomerID As Integer, ByVal newFirstName As String, ByVal newSurname As String, ByVal newAddress As String, ByVal newTown As String, ByVal newCounty As String, ByVal newPostCode As String, ByVal newPhone As String, ByVal newEmail As String)

        CustomerID = newCustomerID
        FirstName = newFirstName
        Surname = newSurname
        Address = newAddress
        Town = newTown
        County = newCounty
        PostCode = newPostCode
        Phone = newPhone
        Email = newEmail

    End Sub




End Class
