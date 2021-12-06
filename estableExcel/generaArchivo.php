<?PHP
// FTP server details
$ftpHost   = '192.168.56.1';
$ftpUsername = 'prueba';
$ftpPassword = 'asd123';

// open an FTP connection
$connId = ftp_connect($ftpHost) or die("Couldn't connect to $ftpHost");

// login to FTP server
$ftpLogin = ftp_login($connId, $ftpUsername, $ftpPassword);

ftp_pasv ($connId, true) ;

ftp_chdir($connId, "EntradaLocal");

$ftp_carpeta_local =  "C:/Users/carlos.jimenez/Downloads/";
rename("C:/Users/carlos.jimenez/Downloads/Archivo_".date('Y-m-d').".xls", "C:/Users/carlos.jimenez/Downloads/Archivo_".date('Y-m-d').".csv");
$ftp_carpeta_remota= "/EntradaLocal/";

$mi_nombredearchivoLocal="Archivo_".date('Y-m-d').".xls";
$mi_nombredearchivoAGuardar="Archivo_".date('Y-m-d').".csv";
$nombre_archivo = $ftp_carpeta_local.$mi_nombredearchivoLocal;
$archivo_destino = $ftp_carpeta_remota.$mi_nombredearchivoAGuardar;
$upload = ftp_put($connId, $archivo_destino, $nombre_archivo, FTP_BINARY);
ftp_close($connId);
?>