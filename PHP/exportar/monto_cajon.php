<?php
    // Obtener el valor del monto enviado por AJAX
    // Configurar cabeceras CORS
    header("Access-Control-Allow-Origin: *"); // Permitir cualquier origen. Puedes cambiar "*" por el dominio específico de tu sitio web.
    header("Access-Control-Allow-Methods: GET, POST, PUT, DELETE"); // Permitir los métodos HTTP que deseas permitir.
    header("Access-Control-Allow-Headers: Content-Type, Authorization"); // Permitir los encabezados HTTP que deseas permitir.

    // Resto de tu código PHP aquí
    $monto = $_GET['monto'];
    // Escribir el valor del monto en un archivo de texto
    if ($monto == "1"){
        $file = 'monto_cajon.txt';
        file_put_contents($file, "");
    }else{
        $cantidad = str_replace(".", ",", $monto);
        $file = 'monto_cajon.txt'; // Ruta del archivo de texto
        file_put_contents($file, $cantidad);
    }
    
    // Devolver una respuesta al cliente
    echo "El monto se ha escrito en el archivo correctamente";
?>
