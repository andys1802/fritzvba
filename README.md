# fritzvba
VBA script for accessing Fritz!Box

Es gibt viele Projekte, die auf die FritzBox zuzugreifen und Funktionen aufzurufen, u.a.

    https://github.com/Kruemelino/FritzBoxTelefon-dingsbums
    http://home.mengelke.de/cgi-bin/webcm?sid=0123456789abcdef

Ich stelle hier trotzdem ein kleines VBA Skript rein, das größtenteils auf den Skripten von Michael Engelke beruht ( http://www.MEngelke.de), dass mir meine Outlook-Kontakte (Exchange) in die FritzBox importiert und das aktuelle Telefonbuch überschreibt. Ja, das geht auch über Google-Konto, ich weiß, ich weiß. Aber ich will Google, 1und1 & Co ja meine Kontaktdaten gar nicht geben.

Evtl. ist es für den ein oder anderen interessant.

Verwendung:

Public Sub ExportContacts()

Dim fb As FritzBox
Dim pbid As String
Dim count As Integer

Set fb = New FritzBox

If fb.Login("1.1.1.1", "password", False) Then
    pbid = fb.getCurrentPhoneBook
    If (fb.uploadPhoneBook(pbid, count)) Then
      MsgBox count & " Kontakt(e) erfolgreich exportiert", vbInformation & vbOKOnly, "Erfolg!"
    End If
    fb.LogOut
End If

End Sub

Viel Spaß!

Lizenz

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program; if not, see http://www.gnu.org/licenses/.
