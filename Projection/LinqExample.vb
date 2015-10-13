Public Class MusicalArtist
    Private _name As String
    Private _genre As String
    Private _latestHit As String
    Private _albums As List(Of Album)

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Genre As String
        Get
            Return _genre
        End Get
        Set(ByVal value As String)
            _genre = value
        End Set
    End Property

    Public Property latestHit As String
        Get
            Return _latestHit
        End Get
        Set(ByVal value As String)
            _latestHit = value
        End Set
    End Property

    Public Property Albums() As List(Of Album)
        Get
            Return _albums
        End Get
        Set(value As List(Of Album))
            _albums = value
        End Set
    End Property
End Class

Public Class Album
    Private _name As String
    Private _year As String

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Year As String
        Get
            Return _year
        End Get
        Set(ByVal value As String)
            _year = value
        End Set
    End Property
End Class

Public Class ArtistViewModel
    Private _artistName As String
    Private _song As String

    Public Property ArtistName As String
        Get
            Return _artistName
        End Get
        Set(value As String)
            _artistName = value
        End Set
    End Property

    Public Property Song As String
        Get
            Return _song
        End Get
        Set(value As String)
            _song = value
        End Set
    End Property
End Class
Public Module LinqExample
    Private Function getMusicalArtists() As List(Of MusicalArtist)
        Return New List(Of MusicalArtist) From {
                        New MusicalArtist() With {
                            .Name = "Adele",
                            .Genre = "Pop",
                            .latestHit = "Someone Like You",
                            .Albums = New List(Of Album)() From {
                                New Album() With {.Name = "21", .Year = "2011"},
                                New Album() With {.Name = "19", .Year = "2008"}
                            }
                        },
                        New MusicalArtist() With {
                            .Name = "Backstreet Boys",
                            .Genre = "Pop",
                            .latestHit = "Larger than Life",
                            .Albums = New List(Of Album)() From {
                                New Album() With {.Name = "Show Me The Meaning", .Year = "2000"},
                                New Album() With {.Name = "I Want It That Way", .Year = "2002"}
                            }
                        },
                        New MusicalArtist() With {
                            .Name = "Shakira",
                            .Genre = "Alternative Rock",
                            .latestHit = "Rabioza",
                            .Albums = New List(Of Album)() From {
                                New Album() With {.Name = "Whenever", .Year = "2009"},
                                New Album() With {.Name = "Wherever", .Year = "2012"}
                            }
                        }
                    }

    End Function

    Public Sub testLinq4()
        Dim artistsDataSource As List(Of MusicalArtist) = getMusicalArtists()

        Dim artistsResult =
            From artist In artistsDataSource
            Select New With {
                Key .Name = artist.Name,
                Key .NumberOfAlbums = artist.Albums.Count}
        For Each artist In artistsResult
            MsgBox(artist.Name)
            MsgBox(artist.NumberOfAlbums)
        Next
    End Sub
    Public Sub testLinq3()
        Dim artistsDataSource As List(Of MusicalArtist) = getMusicalArtists()

        Dim artistResult As IEnumerable(Of ArtistViewModel) =
            From artist In artistsDataSource
            Select New ArtistViewModel With {
                .ArtistName = artist.Name,
                .Song = artist.latestHit}
        For Each artist As ArtistViewModel In artistResult
            MsgBox(artist.ArtistName)
            MsgBox(artist.Song)
        Next
    End Sub

    Public Sub testLinq2()
        Dim artistsDataSource As List(Of MusicalArtist) = getMusicalArtists()
        Dim artistsResults As IEnumerable(Of MusicalArtist) =
            From artist In artistsDataSource
            Select New MusicalArtist With {
                .Name = artist.Name,
                .latestHit = artist.latestHit}
        For Each artist As MusicalArtist In artistsResults
            MsgBox(artist.Name)
            MsgBox(artist.latestHit)
        Next
    End Sub
    Public Sub testLinq()
        Dim musicalArtists As String() = {"Adele", "Backstreet Boys", "Chris"}

        Dim aArtists As IEnumerable(Of String) =
            From artist In musicalArtists
            Where artist.StartsWith("A")
            Select artist

        For Each artist In aArtists
            MsgBox(artist)
        Next
    End Sub
End Module
