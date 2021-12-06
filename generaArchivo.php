<?PHP
include('Net/SFTP.php'); 
/*ini_set('display_errors', 1); 
ini_set('display_startup_errors', 1); 
error_reporting(E_ALL);*/
$strServer = 'ftp.hycite.com';
$intPort = 22;
$strUsername = 'payrollplus-mx';
$strPassword = 'Ayp$QuetOornOcdeb8';

// Instanciamos la clase
$objFtp = new Net_SFTP( $strServer , $intPort );

// Realizamos el logueo
if (!$objFtp ->login( $strUsername , $strPassword )) {
	 exit( 'Login Failed' );
}
$objFtp->chdir('Fortia/Prod/Moves_Newhires'); //Cambia el directorio

$filename = "Archivo_" . date('Y-m-d') . ".csv";

$objFtp->put($filename, $strLocalFile, NET_SFTP_LOCAL_FILE);
$objFtp->disconnect();
//Se elimina el archivo que se realizó de la carpeta de Descargas
$strLocalFile = '/home/jadelrio/Descargas/'.$filename;
unlink($strLocalFile);
die();
?>