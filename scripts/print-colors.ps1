Write-Host "Esta es la primera línea en azul." # Esto se mostrará en blanco en el log de GitHub Actions
Write-Host "Esta es la segunda línea en rojo." # Esto se mostrará en blanco en el log de GitHub Actions
Write-Host "Esta es la tercera línea en amarillo." # Esto se mostrará en blanco en el log de GitHub Actions

# Usamos comandos de workflow para mostrar mensajes con colores en el log de GitHub Actions
Write-Host "::notice title=Línea Azul::Esta es la primera línea que se verá con un fondo azul claro en el log de GitHub Actions."
Write-Host "::warning title=Línea Roja::Esta es la segunda línea que se verá con un fondo amarillo en el log de GitHub Actions (advertencia)."
Write-Host "::error title=Línea Amarilla::Esta es la tercera línea que se verá con un fondo rojo en el log de GitHub Actions (error)."

# Definir códigos ANSI para los colores
$Blue = "`e[34m"
$Red = "`e[31m"
$Yellow = "`e[33m"
$Reset = "`e[0m" # Código para resetear el color a por defecto

# Imprimir las líneas usando los códigos ANSI
Write-Output "${Blue}Esta es la primera línea en azul.${Reset}"
Write-Output "${Red}Esta es la segunda línea en rojo.${Reset}"
Write-Output "${Yellow}Esta es la tercera línea en amarillo.${Reset}"