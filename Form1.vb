Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Data
Imports System.Configuration
Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Net
Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.Runtime.CompilerServices

Public Class Form1
    Private codeSequence As String
    Private lastKeyPressTime As DateTime = DateTime.Now
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = &H2

    Public connectionString As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim configFile As String = "\\scribe\commun\Test informatique\fiche suivi\new\3.1\suivi.CFG"
        If File.Exists(configFile) Then
            ' Charger la configuration du fichier
            Dim configMap As New ExeConfigurationFileMap()
            configMap.ExeConfigFilename = configFile
            Dim config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None)

            ' Récupérer les valeurs du fichier CFG
            Dim server As String = config.AppSettings.Settings("server").Value
            Dim database As String = config.AppSettings.Settings("Database").Value
            Dim uid As String = config.AppSettings.Settings("uid").Value
            Dim pwd As String = config.AppSettings.Settings("pwd").Value

            ' Construire la chaîne de connexion
            connectionString = $"Server={server};Database={database};Uid={uid};Pwd={pwd}"

            ' Afficher la chaîne de connexion dans la zone de texte txtAdress
            TxtAdress.Text = connectionString
            TxtSQL.Text = configFile + " : Connexion réussie"
        Else
            TxtSQL.Text = configFile + " : Connexion échouée !"
            MessageBox.Show("Le fichier CFG n'existe pas.", "Erreur", MessageBoxButtons.OK)
        End If
        TxtAdress.Visible = False
        TxtSQL.Visible = False
        Button1.Visible = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
    End Sub


    Private Sub Form1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
        ' Ajouter chaque caractère de la séquence dans une variable
        codeSequence &= e.KeyChar

        ' Vérifier si la séquence complète "developpermodeon" a été tapée
        If codeSequence.ToLower() = "developpermodeon" Then
            ShowDeveloperElements()
            codeSequence = "" ' Réinitialiser la séquence
        End If

        ' Vérifier si la séquence complète "developpermodeoff" a été tapée
        If codeSequence.ToLower() = "developpermodeoff" Then
            HideDeveloperElements()
            codeSequence = "" ' Réinitialiser la séquence
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Vérifier si aucun caractère n'a été tapé depuis plus de 5 secondes
        If DateTime.Now - lastKeyPressTime > TimeSpan.FromSeconds(5) Then
            codeSequence = "" ' Réinitialiser la séquence
        End If
    End Sub

    Private Sub ShowDeveloperElements()
        ' Afficher les éléments supplémentaires du mode développeur
        TxtAdress.ReadOnly = True
        TxtAdress.Visible = True
        TxtSQL.Visible = True
        TxtSQL.ReadOnly = True

        ' Ajouter ici d'autres éléments à afficher pour le mode développeur
    End Sub

    Private Sub HideDeveloperElements()
        ' Cacher les éléments du mode développeur
        TxtAdress.ReadOnly = False
        TxtAdress.Visible = False
        TxtSQL.Visible = False
        TxtSQL.ReadOnly = False

        ' Ajouter ici d'autres éléments à cacher pour le mode développeur
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox1.Text = "ok"
        Else
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            RadioButton13.Checked = False
            TextBox1.Enabled = True
        ElseIf RadioButton2.Checked = False Then
            TextBox1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        If RadioButton13.Checked = True Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            TextBox1.Text = "non présent"
        Else
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            TextBox2.Text = "ok"
        Else
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            TextBox2.Enabled = True
        ElseIf RadioButton4.Checked = False Then
            TextBox2.Enabled = False
        End If
    End Sub

    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged
        If RadioButton14.Checked = True Then
            TextBox2.Text = "non présent"
        Else
            TextBox2.Text = ""

        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            TextBox3.Text = "ok"
        Else
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked = True Then
            TextBox3.Enabled = True
        ElseIf RadioButton6.Checked = False Then
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        If RadioButton15.Checked = True Then
            TextBox3.Text = "non présent"
        Else
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked = True Then
            TextBox4.Text = "ok"
        Else
            TextBox4.Text = ""

        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If RadioButton8.Checked = True Then
            TextBox4.Enabled = True
        ElseIf RadioButton3.Checked = False Then
            TextBox4.Enabled = False
        End If
    End Sub

    Private Sub RadioButton16_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton16.CheckedChanged
        If RadioButton16.Checked = True Then
            TextBox4.Text = "non présent"
        Else
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If RadioButton9.Checked = True Then
            TextBox5.Text = "ok"
        Else
            TextBox5.Text = ""
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        If RadioButton10.Checked = True Then
            TextBox5.Enabled = True
        Else
            TextBox5.Enabled = False
        End If
    End Sub

    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton17.CheckedChanged
        If RadioButton17.Checked = True Then
            TextBox5.Text = "non présent"
        Else
            TextBox5.Text = ""
        End If
    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged
        If RadioButton11.Checked = True Then
            TextBox6.Text = "ok"
        Else
            TextBox6.Text = ""
        End If
    End Sub

    Private Sub RadioButton12_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton12.CheckedChanged
        If RadioButton12.Checked = True Then
            TextBox6.Enabled = True
        Else
            TextBox6.Enabled = False
        End If
    End Sub

    Private Sub RadioButton18_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton18.CheckedChanged
        If RadioButton18.Checked = True Then
            TextBox6.Text = "non présent"
        Else
            TextBox6.Text = ""
        End If
    End Sub
    Private Sub Checkbox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged, RadioButton9.CheckedChanged, RadioButton10.CheckedChanged
        ' Apparition du bouton "Envoyer"
        If CheckBox1.Checked = True Then
            Button1.Enabled = True
            Button1.Visible = True
        Else
            Button1.Enabled = False
            Button1.Visible = False
        End If
        ' Vérification que tout est bien coché
        If (RadioButton1.Checked Or RadioButton2.Checked Or RadioButton13.Checked) = True AndAlso
(RadioButton3.Checked Or RadioButton4.Checked Or RadioButton14.Checked) = True AndAlso
(RadioButton5.Checked Or RadioButton6.Checked Or RadioButton15.Checked) = True AndAlso
(RadioButton7.Checked Or RadioButton8.Checked Or RadioButton16.Checked) = True AndAlso
(RadioButton9.Checked Or RadioButton10.Checked Or RadioButton17.Checked) = True AndAlso
(RadioButton11.Checked Or RadioButton12.Checked Or RadioButton18.Checked) = True Then

            CheckBox1.Enabled = True
            CheckBox1.Visible = True
        Else
            CheckBox1.Enabled = False
        End If

        ' Vérification que les TextBox sont remplies
        If TextBox1.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant l'unité centrale ou le client léger.")
            CheckBox1.Checked = False
            TextBox1.Enabled = True
        ElseIf (TextBox1.Text = "ok" Or TextBox1.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox1.Enabled = False
        End If
        If TextBox2.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant l'écran.")
            CheckBox1.Checked = False
            TextBox2.Enabled = True
        ElseIf (TextBox2.Text = "ok" Or TextBox2.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox2.Enabled = False
        End If
        If TextBox3.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant le clavier et la souris.")
            CheckBox1.Checked = False
            TextBox3.Enabled = True
        ElseIf (TextBox3.Text = "ok" Or TextBox3.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox3.Enabled = False
        End If
        If TextBox4.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant les câbles (internet, alimentation, connectique).")
            CheckBox1.Checked = False
            TextBox4.Enabled = True
        ElseIf (TextBox4.Text = "ok" Or TextBox4.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox4.Enabled = False
        End If
        If TextBox5.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant les câbles (internet, alimentation, connectique).")
            CheckBox1.Checked = False
            TextBox5.Enabled = True
        ElseIf (TextBox5.Text = "ok" Or TextBox5.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox5.Enabled = False
        End If
        If TextBox6.Text = "" AndAlso CheckBox1.Checked = True Then
            MsgBox("Merci de remplir la partie 'observation' concernant les câbles (internet, alimentation, connectique).")
            CheckBox1.Checked = False
            TextBox6.Enabled = True
        ElseIf (TextBox6.Text = "ok" Or TextBox6.Text <> "") AndAlso CheckBox1.Checked = True Then
            TextBox6.Enabled = False
        End If

        ' vérification que le professeur a bien saisi le nom de la salle 
        If TextBox7.Text = "" AndAlso CheckBox1.Checked = True Then
            CheckBox1.Checked = False
            MsgBox("Merci d'indiquer le nom de la salle")
        ElseIf TextBox7.Text <> "" AndAlso CheckBox1.Checked = True Then
            TextBox7.Enabled = False
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim deuxiemeForm As New Form2()

        deuxiemeForm.Show()
    End Sub

    ' Méthode pour créer un nouveau ticket dans GLPI
    ' Méthode pour lire le jeton API depuis le fichier de configuration
    Private Function LireJetonAPI() As String
        Try
            ' Chemin vers le fichier de configuration externe
            Dim configFile As String = "C:\Users\Administrateur\Desktop\Fiche enseignante version Romain.exe.config"
            MessageBox.Show("Chemin du fichier : " & Path.GetFullPath(configFile)) ' Afficher le chemin complet

            ' Vérifier si le fichier de configuration existe
            If File.Exists(configFile) Then
                ' Lire le jeton API depuis le fichier
                Dim configMap As New ExeConfigurationFileMap()
                configMap.ExeConfigFilename = configFile
                Dim config As Configuration = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None)

                Dim apiKey As String = config.AppSettings.Settings("config").Value
                Dim configSetting As KeyValueConfigurationElement = config.AppSettings.Settings("config")
                If configSetting IsNot Nothing Then
                    apiKey = configSetting.Value
                Else
                    MessageBox.Show("La clé 'config' n'existe pas dans les paramètres de l'application.")
                End If
            Else
                MessageBox.Show("Fichier de configuration introuvable." & configFile)
                Return ""
            End If
        Catch ex As Exception
            ' Gérer les erreurs en conséquence
            MessageBox.Show("Une erreur s'est produite lors de la lecture du fichier de configuration : " & ex.Message)
            Return ""
        End Try
    End Function

    ' Méthode pour créer un nouveau ticket dans GLPI
    Private Sub CreerTicketGLPI()
        Try
            ' URL de l'API GLPI pour créer un ticket
            Dim apiUrl As String = "http://VOTRE_ADRESSE_GLPI/apirest.php/initSession"

            ' Récupérer le jeton API depuis le fichier de configuration
            Dim apiToken As String = LireJetonAPI()

            ' Vérifier si le jeton API a été correctement récupéré
            If String.IsNullOrEmpty(apiToken) Then
                Exit Sub ' Quitter la méthode si le jeton est vide ou invalide
            End If

            ' Création du corps de la requête en JSON
            Dim jsonBody As String = "{""input"": {""name"": ""NOM_DU_TICKET"", ""content"": ""CONTENU_DU_TICKET"", ""status"": 1, ""urgency"": 3, ""impact"": 3, ""type"": 2}}"

            ' Convertir le corps de la requête en bytes
            Dim data As Byte() = Encoding.UTF8.GetBytes(jsonBody)

            ' Configuration de la requête
            Dim request As HttpWebRequest = DirectCast(WebRequest.Create(apiUrl), HttpWebRequest)
            request.Method = "POST"
            request.ContentType = "application/json"
            request.ContentLength = data.Length

            ' Ajout du jeton API dans l'en-tête de la requête
            request.Headers("Authorization") = "user_token " & apiToken

            ' Écrire le corps de la requête dans le flux de sortie de la requête
            Using stream As Stream = request.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using

            ' Exécution de la requête et récupération de la réponse
            Using response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
                Using streamReader As New StreamReader(response.GetResponseStream())
                    Dim responseContent As String = streamReader.ReadToEnd()
                    ' Traiter la réponse si nécessaire
                    ' Ici, vous pouvez analyser la réponse JSON pour obtenir des informations sur le ticket créé
                    MessageBox.Show("Ticket créé avec succès !")
                End Using
            End Using
        Catch ex As Exception
            ' Gérer les erreurs en conséquence
            MessageBox.Show("Une erreur s'est produite : " & ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CreerTicketGLPI()
    End Sub
End Class


