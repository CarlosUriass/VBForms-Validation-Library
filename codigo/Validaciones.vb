Public Class Validaciones

    Public Shared Function validacion_vacio(ByVal texto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Funci�n: validacion_vacio
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si un texto dado est� vac�o o no. Si el texto est� vac�o,
        '   muestra un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Par�metros:
        '   texto (String): El texto que se va a verificar si est� vac�o o no.
        '
        ' Devuelve:
        '   Boolean: True si el texto est� vac�o, False si no lo est�.
        '-------------------------------------------------------------------------------

        ' Verifica si el texto est� vac�o
        If texto = "" Then
            ' Si el texto est� vac�o, muestra un mensaje de advertencia
            MsgBox("No puede haber campos vac�os", MsgBoxStyle.OkOnly)
            Return True
        Else
            ' Si el texto no est� vac�o, devuelve False
            Return False
        End If
    End Function



    Public Shared Function validarLongitud(ByVal texto As String, ByVal longitud As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Funci�n: validarLongitud
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si un texto dado tiene una longitud especificada. Si la longitud
        '   del texto no coincide con la longitud especificada, muestra un mensaje de advertencia y
        '   devuelve True; de lo contrario, devuelve False.
        '
        ' Par�metros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   longitud (Integer): La longitud requerida del texto que se va a verificar.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto no es igual a la longitud especificada,
        '            False si lo es.
        '-------------------------------------------------------------------------------


        ' Verifica si la longitud del texto es diferente de la longitud especificada
        If Len(texto) <> longitud Then
            ' Si la longitud del texto no es exactamente la longitud especificada, devuelve True y muestra un mensaje de advertencia
            MsgBox("El campo debe tener exactamente " & longitud & " caracteres de longitud", MsgBoxStyle.OkOnly)
            Return True
        Else
            ' Si la longitud coincide con la longitud especificada, devuelve False
            Return False
        End If
    End Function

    '-------------------------------------------------------------------------------
    ' Funci�n: soloNumeros
    '-------------------------------------------------------------------------------
    ' Descripci�n:
    '   Esta funci�n verifica si un texto dado contiene solo caracteres num�ricos. 
    '   Si encuentra alg�n car�cter que no sea un n�mero, muestra un mensaje de advertencia
    '   y devuelve True; de lo contrario, devuelve False.
    '
    ' Par�metros:
    '   texto (String): El texto que se va a verificar para asegurar que contenga solo n�meros.
    '
    ' Devuelve:
    '   Boolean: True si el texto contiene caracteres que no son n�meros, False si todos los caracteres son n�meros.
    '-------------------------------------------------------------------------------

    Public Shared Function soloNumeros(ByVal texto As String) As Boolean
        ' Validar que solo sean n�meros
        For i As Integer = 1 To Len(texto)
            If Not Char.IsDigit(texto(i - 1)) Then
                ' Si encuentra un car�cter que no es un n�mero, muestra un mensaje de advertencia y devuelve True
                MsgBox("El campo solo debe contener n�meros", MsgBoxStyle.OkOnly)
                Return True
                Exit For ' Salir del bucle luego de encontrar el primer car�cter no num�rico
            End If
        Next
        Return False
    End Function

    Public Shared Function caracteresEspeciales(ByVal texto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Funci�n: caracteresEspeciales
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si un texto contiene caracteres especiales. Si encuentra
        '   alg�n car�cter especial, muestra un mensaje de advertencia y devuelve True; de lo
        '   contrario, devuelve False.
        '
        ' Par�metros:
        '   texto (String): El texto que se va a verificar en busca de caracteres especiales.
        '
        ' Devuelve:
        '   Boolean: True si el texto contiene caracteres especiales, False si no los contiene.
        '
        ' Caracteres Especiales en la Tabla ASCII:
        '   32  (space)
        '   33  !
        '   34  "
        '   35  #
        '   36  $
        '   37  %
        '   38  &
        '   39  '
        '   40  (
        '   41  )
        '   42  *
        '   43  +
        '   44  ,
        '   47  /
        '   58  :
        '   59  ;
        '   60  <
        '   61  =
        '   62  >
        '   63  ?
        '   64  @
        '   91  [
        '   92  \
        '   93  ]
        '   94  ^
        '   123 {
        '   124 |
        '   125 }
        '   126 ~

        'Nota: El 45 (-), 46(.) y el 95(_) fueron removidos para tener compatibilidad con la validaci�n de correo electronico
        '-------------------------------------------------------------------------------

        For i As Integer = 1 To Len(texto)
            Select Case Asc(texto(i - 1))
                Case 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 47, 58, 59, 60, 61, 62, 63, 64, 91, 92, 93, 94, 123, 124, 125, 126
                    MsgBox("El campo no puede contener caracteres especiales", MsgBoxStyle.OkOnly)
                    Return True
                    Exit For
            End Select
        Next

        Return False
    End Function


    Public Shared Function validarDinero(ByVal monto As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Funci�n: validarDinero
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si una cadena dada representa un monto de dinero v�lido.
        '   Verifica si la cadena contiene solo n�meros y opcionalmente un punto decimal,
        '   lo que indica la separaci�n entre la parte entera y la parte decimal del monto.
        '
        ' Par�metros:
        '   monto (String): La cadena que representa el monto de dinero que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si la cadena no cumple con los criterios de validaci�n, False si
        '            cumple con los criterios y es considerada un monto de dinero v�lido.
        '
        ' Dependencias:
        '   Esta funci�n hace uso de la funci�n soloNumeros previamente definida para verificar
        '   si una cadena contiene solo caracteres num�ricos.
        '-------------------------------------------------------------------------------
        Dim punto_posicion As Byte

        ' Buscar la posici�n del punto decimal en la cadena
        For i As Integer = 1 To Len(monto)
            If monto(i - 1) = "." Then
                punto_posicion = i
                Exit For
            End If
        Next

        ' Validar que solo sean n�meros antes del punto decimal
        If soloNumeros(Mid(monto, 1, punto_posicion - 1)) Then
            Return True
        End If

        ' Validar que solo sean n�meros despu�s del punto decimal (si existe)
        If punto_posicion > 0 AndAlso soloNumeros(Mid(monto, punto_posicion + 1)) Then
            Return True
        End If

        ' Si todas las validaciones pasan, entonces el monto de dinero es v�lido
        Return False
    End Function



    Public Shared Function mayorquecero(ByVal numero As Integer) As Boolean
        '-------------------------------------------------------------------------------
        ' Funci�n: mayorquecero
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si un n�mero entero es mayor que cero. Si el n�mero es
        '   menor que cero, muestra un mensaje de advertencia y devuelve True; de lo contrario,
        '   devuelve False.
        '
        ' Par�metros:
        '   numero (Integer): El n�mero entero que se va a verificar si es mayor que cero.
        '
        ' Devuelve:
        '   Boolean: True si el n�mero es menor que cero, False si es mayor o igual a cero.
        '-------------------------------------------------------------------------------

        If numero < 0 Then
            MsgBox("El campo no puede ser menor que 0", MsgBoxStyle.OkOnly)
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function longitudMaxima(ByVal texto As String, maximo As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Funci�n: longitudMaxima
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si la longitud de un texto dado excede un l�mite m�ximo
        '   especificado. Si la longitud del texto es mayor que el l�mite m�ximo, muestra
        '   un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Par�metros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   maximo (Integer): La longitud m�xima permitida para el texto.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto excede el l�mite m�ximo, False si no lo hace.
        '-------------------------------------------------------------------------------

        If Len(texto) > maximo Then
            MsgBox("El campo excede los caracteres permitidos", MsgBoxStyle.OkOnly)
            Return True
        End If

        Return False


    End Function

    Public Shared Function longitudMinima(ByVal texto As String, minimo As Integer) As Boolean

        '-------------------------------------------------------------------------------
        ' Funci�n: longitudMinima
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si la longitud de un texto dado es menor que un l�mite
        '   m�nimo especificado. Si la longitud del texto es menor que el l�mite m�nimo,
        '   muestra un mensaje de advertencia y devuelve True; de lo contrario, devuelve False.
        '
        ' Par�metros:
        '   texto (String): El texto cuya longitud se va a verificar.
        '   minimo (Integer): La longitud m�nima requerida para el texto.
        '
        ' Devuelve:
        '   Boolean: True si la longitud del texto es menor que el l�mite m�nimo, False si no lo es.
        '-------------------------------------------------------------------------------


        If Len(texto) < minimo Then
            MsgBox("El campo debe tener al menos " & minimo & " caracteres", MsgBoxStyle.OkOnly)
            Return True
        End If


        Return False
    End Function


    Public Shared Function numerotelefonico(ByVal numero As String) As Boolean
        '-------------------------------------------------------------------------------
        ' Funci�n: numerotelefonico
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si un n�mero telef�nico dado cumple con los criterios
        '   de validaci�n establecidos. Primero, verifica si el n�mero contiene solo
        '   caracteres num�ricos. Luego, verifica si el n�mero tiene una longitud de
        '   10 d�gitos, que es la longitud est�ndar para n�meros telef�nicos en M�xico.
        '
        ' Par�metros:
        '   numero (String): El n�mero telef�nico que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si el n�mero telef�nico cumple con los criterios de validaci�n,
        '            False si no los cumple.
        '
        ' Dependencias:
        '   Esta funci�n utiliza las siguientes funciones de validaci�n definidas previamente:
        '   - soloNumeros: Verifica si una cadena contiene solo caracteres num�ricos.
        '   - validarLongitud: Verifica si una cadena tiene una longitud espec�fica.
        '-------------------------------------------------------------------------------

        'Primero validar que sean puros numeros

        If soloNumeros(numero) = True Then
            Return True
        End If

        'Validar longitud

        If validarLongitud(numero, 10) = True Then
            Return True
        End If


        Return False

    End Function

    Public Shared Function correoelectronico(ByVal email As String) As Boolean

        '-------------------------------------------------------------------------------
        ' Funci�n: correoelectronico
        '-------------------------------------------------------------------------------
        ' Descripci�n:
        '   Esta funci�n verifica si una cadena dada cumple con los criterios para ser
        '   considerada como una direcci�n de correo electr�nico v�lida. Verifica si la
        '   cadena tiene un formato adecuado, incluyendo la presencia de una �nica arroba,
        '   longitud del nombre de usuario y del dominio dentro de los l�mites permitidos,
        '   ausencia de caracteres especiales en el nombre de usuario, y si el dominio
        '   tiene una extensi�n de dominio (TLD) v�lida (.com, .gob, .org, .net).
        '
        ' Par�metros:
        '   email (String): La cadena que representa la direcci�n de correo electr�nico
        '                   que se va a validar.
        '
        ' Devuelve:
        '   Boolean: True si la cadena no cumple con los criterios de validaci�n, False si
        '            cumple con los criterios y es considerada una direcci�n de correo
        '            electr�nico v�lida.
        '
        ' Dependencias:
        '   Esta funci�n hace uso de las siguientes funciones de validaci�n previamente
        '   definidas:
        '   - longitudMaxima: Verifica si una cadena excede una longitud m�xima especificada.
        '   - longitudMinima: Verifica si una cadena tiene una longitud m�nima especificada.
        '   - caracteresEspeciales: Verifica si una cadena contiene caracteres especiales.
        '-------------------------------------------------------------------------------


        Dim arroba_posc As Integer
        Dim contador_arroba As Integer
        contador_arroba = 0

        ' Contar la cantidad de arrobas en el correo electr�nico
        For i As Integer = 1 To Len(email)
            If email(i - 1) = "@" Then
                arroba_posc = i
                contador_arroba += 1
            End If
        Next

        If contador_arroba > 1 Then
            MsgBox("Solo debe existir una arroba en el email", MsgBoxStyle.OkOnly)
            Return True
        End If

        If contador_arroba = 0 Then
            MsgBox("El correo debe contener al menos un arroba", MsgBoxStyle.OkOnly)
            Return True
        End If

        'Generar nombre de usuario
        Dim nombre_usuario As String
        nombre_usuario = Microsoft.VisualBasic.Left(email, arroba_posc - 1)

        'Validar si el nombre de usuario excede los 64 caracteres

        If longitudMaxima(nombre_usuario, 64) = True Then
            Return True
        End If

        'Validar que el nombre de usuario sea mayor que 1

        If longitudMinima(nombre_usuario, 2) = True Then
            Return True
        End If

        'Generar dominio
        Dim dominio As String
        dominio = Mid(email, arroba_posc + 1)

        'Validar que el dominio no exceda los 255 caracteres

        If longitudMaxima(dominio, 64) = True Then
            Return True
        End If

        'Validar que el nombre de usuario no tenga caracteres especiales

        If caracteresEspeciales(nombre_usuario) = True Then
            Return True
        End If

        ' Validar que el dominio tenga una extensi�n de dominio (TLD) v�lida
        If Not (dominio.EndsWith(".com") Or dominio.EndsWith(".gob") Or dominio.EndsWith(".org") Or dominio.EndsWith(".net")) Then
            MsgBox("El dominio del correo electr�nico no es v�lido.", MsgBoxStyle.OkOnly)
            Return True
        End If


        Return False

    End Function

End Class
