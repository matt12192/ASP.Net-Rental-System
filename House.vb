Public Class House


    'Attributes which define the House information

    Private intHouseReference As Integer
    Private intCustomerID As Integer
    Private strAddress As String
    Private strTown As String
    Private strCounty As String
    Private strPostCode As String
    Private strType As String
    Private strNoOfBedrooms As String
    Private strNoOfReceiptionRooms As String
    Private strGarage As String
    Private strRentalCosts As String
    Private strLeaseLength As String
    Private strStatus As String

    'Properties to set a attribute values and return a value from an attribute

    Public Property HouseReference As Integer

        Get

            Return intHouseReference

        End Get

        Set(ByVal value As Integer)

            intHouseReference = value

        End Set

    End Property

    Public Property CustomerID As Integer

        Get

            Return intCustomerID

        End Get

        Set(ByVal value As Integer)

            intCustomerID = value

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

    

    Public Property Type As String

        Get

            Return strType

        End Get

        Set(ByVal value As String)

            strType = value

        End Set

    End Property

    Public Property NoOfBedrooms As String

        Get

            Return strNoOfBedrooms

        End Get

        Set(ByVal value As String)

            strNoOfBedrooms = value

        End Set

    End Property

    Public Property NoOfReceiptionRooms As String

        Get

            Return strNoOfReceiptionRooms

        End Get

        Set(ByVal value As String)

            strNoOfReceiptionRooms = value

        End Set

    End Property

    Public Property Garage As String

        Get

            Return strGarage

        End Get

        Set(ByVal value As String)

            strGarage = value

        End Set

    End Property

    Public Property RentalCosts As String

        Get

            Return strRentalCosts

        End Get

        Set(ByVal value As String)

            strRentalCosts = value

        End Set

    End Property

    Public Property LeaseLength As String

        Get

            Return strLeaseLength

        End Get

        Set(ByVal value As String)

            strLeaseLength = value

        End Set

    End Property

    Public Property Status As String

        Get

            Return strStatus

        End Get

        Set(ByVal value As String)

            strStatus = value

        End Set

    End Property

    'Default Constructor

    Sub New()

        'Sets default values for each of the properties highlighted above

        HouseReference = 1234
        CustomerID = 1234
        Address = "dftAddress"
        Town = "dftTown"
        County = "dftCounty"
        PostCode = "dftPostCode"
        Type = "dftType"
        NoOfBedrooms = "dftNoOfBedrooms"
        NoOfReceiptionRooms = "dftNoOfReceptionRooms"
        Garage = "dftGarage"
        RentalCosts = "dftRentalCost"
        LeaseLength = "dftLeaseLength"
        Status = "dftStatus"



    End Sub

    'Overloaded Constructor 

    'Allows new House values to be inserted in from external Class code

    Sub New(ByVal newHouseReference As Integer, ByVal newCustomerID As Integer, ByVal newAddress As String, ByVal newTown As String, ByVal newCounty As String, ByVal newPostCode As String, ByVal newType As String, _
            ByVal newNoOfBedrooms As String, ByVal newNoOfReceiptionRooms As String, ByVal newGarage As String, ByVal newRentalCosts As String, ByVal newLeaseLength As String, ByVal newStatus As String)

        HouseReference = newHouseReference
        CustomerID = newCustomerID
        Address = newAddress
        Town = newTown
        County = newCounty
        PostCode = newPostCode
        Type = newType
        NoOfBedrooms = newNoOfBedrooms
        NoOfReceiptionRooms = newNoOfReceiptionRooms
        Garage = newGarage
        RentalCosts = newRentalCosts
        LeaseLength = newLeaseLength
        Status = newStatus

    End Sub




End Class
