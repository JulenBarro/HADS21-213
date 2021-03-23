Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient

Public Class LogicaNegocio
    Public Function enviarEmailRegistro(ByVal emailto As String, ByVal codigo As String) As Boolean
        Try
            'Direccion de origen
            Dim from_address As New MailAddress("PruebasEmail4321@gmail.com", "Email Pruebas")
            'Direccion de destino
            Dim to_address As New MailAddress(emailto)
            'Password de la cuenta
            Dim from_pass As String = 
            'Objeto para el cliente smtp
            Dim smtp As New SmtpClient
            'Host en este caso el servidor de gmail
            smtp.Host = "smtp.gmail.com"
            'Puerto
            smtp.Port = 587
            'SSL activado para que se manden correos de manera segura
            smtp.EnableSsl = True
            'No usar los credenciales por defecto ya que si no no funciona
            smtp.UseDefaultCredentials = False
            'Credenciales
            smtp.Credentials = New System.Net.NetworkCredential(from_address.Address, from_pass)
            'Creamos el mensaje con los parametros de origen y destino
            Dim message As New MailMessage(from_address, to_address)
            'A�adimos el asunto
            message.Subject = "subject"
            'Concatenamos el cuerpo del mensaje a la plantilla
            message.Body = "<html><head></head><body>" + "http://hads21-212.azurewebsites.net/Confirmar.aspx?mbr=" + emailto + "&numconf=" + codigo + "</body></html>"
            'Definimos el cuerpo como html para poder escojer mejor como lo mandamos
            message.IsBodyHtml = True
            'Se envia el correo
            smtp.Send(message)
        Catch e As Exception
            Return False
        End Try
        Return True
    End Function

    Public Function envioCodigoCambioContrase�a(ByVal emailto As String, ByVal codigo As String) As Boolean
        Try
            'Direccion de origen
            Dim from_address As New MailAddress("PruebasEmail4321@gmail.com", "Email Pruebas")
            'Direccion de destino
            Dim to_address As New MailAddress(emailto)
            'Password de la cuenta
            Dim from_pass As String = 
            'Objeto para el cliente smtp
            Dim smtp As New SmtpClient
            'Host en este caso el servidor de gmail
            smtp.Host = "smtp.gmail.com"
            'Puerto
            smtp.Port = 587
            'SSL activado para que se manden correos de manera segura
            smtp.EnableSsl = True
            'No usar los credenciales por defecto ya que si no no funciona
            smtp.UseDefaultCredentials = False
            'Credenciales
            smtp.Credentials = New System.Net.NetworkCredential(from_address.Address, from_pass)
            'Creamos el mensaje con los parametros de origen y destino
            Dim message As New MailMessage(from_address, to_address)
            'A�adimos el asunto
            message.Subject = "subject"
            'Concatenamos el cuerpo del mensaje a la plantilla
            message.Body = "<html><head></head><body>" + "El c�digo para el cambio de contrase�a es el siguiente: " + codigo + "</body></html>"
            'Definimos el cuerpo como html para poder escojer mejor como lo mandamos
            message.IsBodyHtml = True
            'Se envia el correo
            smtp.Send(message)
        Catch e As Exception
            Return False
        End Try
        Return True
    End Function

    Public Function conectar() As String
        Dim aux As String = AccesoDatos.accesodatosSQL.conectar()
        Return aux
    End Function

    Public Function cerrarconexion() As String
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return "Conexion Cerrada"
    End Function

    Public Function insertar(ByVal email As String, ByVal nombre As String, ByVal apellidos As String, ByVal numbconfir As Integer, ByVal confirmado As Boolean, ByVal tipo As String, ByVal pass As String, ByVal codpass As Integer) As String
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux As String = AccesoDatos.accesodatosSQL.insertar(email, nombre, apellidos, numbconfir, confirmado, tipo, pass, codpass)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return (aux)
    End Function

    Public Function emailycodigoCorrectos(ByVal emailto As String, ByVal codigo As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.emailycodigoCorrectos(emailto, codigo)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function cambioConfirmado(ByVal email As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.cambioConfirmado(email)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function existeEmailConfirmado(ByVal email As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.existeEmailConfirmado(email)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function codigoContrase�a(ByVal email As String, ByVal codigo As Integer) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.codigoContrase�a(email, codigo)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function codigoContrase�aGet(ByVal emailto As String) As (numero As Integer, confir As Boolean)
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.codigoContrase�aGet(emailto)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function cambiarContrase�a(ByVal email As String, ByVal pass As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.cambiarContrase�a(email, pass)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function existeEmail(ByVal emailto As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.existeEmail(emailto)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function emailContrase�a(ByVal emailto As String, ByVal pass As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.emailContrase�a(emailto, pass)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function asignaturasAlumno(ByVal email As String) As (ds As DataSet, da As SqlDataAdapter)
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.asignaturasAlumno(email)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function datosParaImportXMLProfesor(ByVal email As String) As (ds As DataSet, da As SqlDataAdapter)
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.datosParaImportXMLProfesor(email)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function asignaturasProfe(ByVal email As String) As (ds As DataSet, da As SqlDataAdapter)
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.asignaturasProfe(email)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    Public Function asignaturasAuxProfesor(ByVal codigo As String) As (ds As DataSet, da As SqlDataAdapter)
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.asignaturasAuxProfesor(codigo)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

    'Devuelve true si es alumno y false si es profesor
    Public Function esAlumno(ByVal emailto As String) As Boolean
        AccesoDatos.accesodatosSQL.conectar()
        Dim aux = AccesoDatos.accesodatosSQL.esAlumno(emailto)
        AccesoDatos.accesodatosSQL.cerrarconexion()
        Return aux
    End Function

End Class

