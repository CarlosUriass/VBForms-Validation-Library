# VBForms-Validation-Library
ValidacionesVB es una biblioteca de funciones en Visual Basic .NET para realizar diversas validaciones comunes en aplicaciones de escritorio y web. Estas funciones pueden ser utilizadas para validar campos de entrada de datos, como texto, números, direcciones de correo electrónico, números telefónicos, entre otros. 

Diseñada principalmente para su uso en **MICROSOFT FORMS**  y escrita en su totalidad en el lenguaje Visual Basic.


## Funciones disponibles

La biblioteca proporciona las siguientes funciones de validación:

- `validacion_vacio`: Verifica si un texto está vacío.
- `validarLongitud`: Verifica si un texto tiene una longitud específica.
- `soloNumeros`: Verifica si un texto contiene solo caracteres numéricos.
- `caracteresEspeciales`: Verifica si un texto contiene caracteres especiales.
- `validarDinero`: Verifica si una cadena representa un monto de dinero válido.
- `mayorquecero`: Verifica si un número entero es mayor que cero.
- `longitudMaxima`: Verifica si la longitud de un texto excede un límite máximo especificado.
- `longitudMinima`: Verifica si la longitud de un texto es menor que un límite mínimo especificado.
- `numerotelefonico`: Verifica si un número telefónico dado cumple con los criterios de validación.
- `correoelectronico`: Verifica si una cadena dada cumple con los criterios para ser considerada como una dirección de correo electrónico válida.

## Uso

Para utilizar esta biblioteca en tu proyecto, simplemente incluye el archivo `Validaciones.vb` en tu proyecto de Visual Basic .NET y llame a las funciones según sea necesario. Puedes llamar a las funciones directamente utilizando el nombre de la clase `Validaciones`, ya que todas las funciones están definidas como compartidas (shared).



```vb
'Propuesta de caso de uso

If Validaciones.validacion_vacio(texto) = True Then
    return ' En caso de que la llamada a la función devuelva true, significa que el texto no cumplió la validacion propuesta -vease mas a fondo en el archivo validaciones.vb- Return "cancelaria" el evento de algun button tipo 'submit'
End If
```


## Contribuir
¡Siéntete libre de contribuir a esta biblioteca! Si encuentras algún error o deseas agregar nuevas funciones de validación, simplemente realiza un fork del repositorio, implementa tus cambios y envía un pull request. ¡Todas las contribuciones son bienvenidas!

## Licencia
Este proyecto está bajo la Licencia MIT. Consulta el archivo LICENSE para más detalles.