<?PHP
include('Net/SFTP.php'); 
/*ini_set('display_errors', 1); 
ini_set('display_startup_errors', 1); 
error_reporting(E_ALL);*/


 /**
  * Establezco la ruta de nuestro sitio de pruebas en el
  * include_path para que la librería pueda incluir ella
  * misma sus archivos necesario
  */
set_include_path(implode(PATH_SEPARATOR, array(
	realpath(dirname(__FILE__) . '/phpseclib'),
	get_include_path(),
)));

// Este bloque de 4 líneas no requiere explicación
$strServer = 'ftp.hycite.com';
$intPort = 22;
$strUsername = 'payrollplus-mx';
$strPassword = 'Ayp$QuetOornOcdeb8';

/**
* Estos serían los archivos local y remoto con los que
* vamos a trabajar
*/
/*$strLocalFile = 'archivo.csv';
$strRemoteFile = 'Prueba JA -1.csv';*/

// Instanciamos la clase
$objFtp = new Net_SFTP( $strServer , $intPort );

// Realizamos el logeo
if (!$objFtp ->login( $strUsername , $strPassword )) {
	 exit( 'Login Failed' );
}
//$objFtp->chdir('Fortia/Prod/Pendientes'); // open directory 'test'
$objFtp->chdir('Fortia/Prod/Moves_Newhires'); // open directory 'test'




/*foreach($objFtp->nlist() as $a){
	echo $a."</br>";
}
die();*/









//die();
/*header("Content-type: application/x-msdownload");
header("Content-Disposition: attachment; filename=Archivo_".date('Y-m-d').".xls");*/
require_once 'PHPExcel/Classes/PHPExcel.php';
// FTP server details

$contador=0;
$nombreArchivoArray = array();
$tabla="";

$delimiter = ",";
//create a file pointer
$f = fopen('php://memory', 'w');

$fields = array('CLA_TIPO_MOV','FECHA_MOV','CLA_TRAB','CLA_EMPRESA','AP_PATERNO','AP_MATERNO',
'NOM_TRAB','CLA_DEPTO','CLA_UBICACION','CLA_UBICACION_PAGO','CLA_CENTRO_COSTO','CLA_PUESTO',
'CLA_CLASIFICACION','CLA_TAB_PRE','CLA_PERIODO','CLA_REG_IMSS','TIPO_SALARIO','SUELDO_DIA',
'SUELDO_MENSUAL','FECHA_ING','FECHA_ING_GRUPO','TIPO_CONTRATO','CLA_FORMA_PAGO','CLA_BANCO',
'CTA_BANCO','CLABE_INTERBANCARIA','RH_DET_POSICION_PLANTILLA','LUGAR_NAC','FECHA_NAC','SEXO',
'SIND','CTA_CORREO','EDO_CIVIL','RFC','NUM_IMSS','CURP',
'CALLE','COLONIA','CP','CIUDAD','ESTADO','CLA_DELEGACION',
'TELEFONO','NACIONALIDAD','CLA_RAZON_SOCIAL','DIAS_CONT','CLA_TAB_SUE','NIV_TAB_SUE',
'Global_ID','FECHA_SAL','FECHA_PER_INFO');
fputcsv($f, $fields, $delimiter);

//SE COMPARA CONTRA LA BASE DE DATOS
$serverName = "10.11.0.34";
$connectionInfo = array( "Database"=>"Royal", "UID"=>"carlosj", "PWD"=>"martha01");
$conn = sqlsrv_connect( $serverName, $connectionInfo);

$localFilePath  = "/var/www/html/Descargas/";
foreach($objFtp->nlist() as $a){
	if($a!='.' && $a!='..'){
		$nombreArchivo = $a;
		$nombreArchivoArray[$contador] = $a;
		$nombreTemp = explode(".", $nombreArchivo);
		$remoteFilePath = '/Fortia/Prod/Pendientes/'.$nombreArchivo;
		if($objFtp->get($a, '/var/www/html/FTP/Descargas/'.$a)){
			//if(ftp_get($connection, "/home/jadelrio/Descargas/".$a, $remoteFilePath, FTP_BINARY)){
			
			//*******************************************************************************//
			//**************************SE PROCESA EL ARCHIVO********************************//
			//*******************************************************************************//
			$archivo = '/var/www/html/FTP/Descargas/'.$a;
			/*$inputFileType = PHPExcel_IOFactory::identify('/var/www/html/FTP/Descargas/'.$a);
			echo "inputFileType->".$inputFileType;
			die();*/
			$objReader = PHPExcel_IOFactory::createReader('CSV');
			$objPHPExcel = $objReader->load($archivo);
			$sheet = $objPHPExcel->getSheet(0); 
			$highestRow = $sheet->getHighestRow();
			$highestColumn = $sheet->getHighestColumn();

			$num=0;
			$numTemp = 0;
			$contador1=0;
			$arrConceptos = array();
			$arrExcel = array();
			$column1=65; //A
			$column2=66; //B
			$column3=67; //C
			$column4=68; //D
			$column5=69; //E
			$column6=70; //F
			$column7=71; //G
			$column8=72; //H
			$column9=73; //I
			$column10=74; //J
			$column11=75; //K
			$column12=76; //L
			$column13=77; //M
			$column14=78; //N
			$column15=79; //O
			$column16=80; //P
			$column17=81; //Q
			$column18=82; //R
			$column19=83; //S
			$column20=84; //T
			$column21=85; //U
			$column22=86; //V
			$column23=87; //W
			$column24=88; //X
			$column25=89; //Y
			$column26=90; //Z
			$controlInicio = 0;

			if($contador==0){
				$tabla = "<table>";
				$tabla.= "<tr>
					<td>CLA_TIPO_MOV</td><td>FECHA_MOV</td><td>CLA_TRAB</td><td>CLA_EMPRESA</td><td>AP_PATERNO</td><td>AP_MATERNO</td>
					<td>NOM_TRAB</td><td>CLA_DEPTO</td><td>CLA_UBICACION</td><td>CLA_UBICACION_PAGO</td><td>CLA_CENTRO_COSTO</td><td>CLA_PUESTO</td>
					<td>CLA_CLASIFICACION</td><td>CLA_TAB_PRE</td><td>CLA_PERIODO</td><td>CLA_REG_IMSS</td><td>TIPO_SALARIO</td><td>SUELDO_DIA</td>
					<td>SUELDO_MENSUAL</td><td>FECHA_ING</td><td>FECHA_ING_GRUPO</td><td>TIPO_CONTRATO</td><td>CLA_FORMA_PAGO</td><td>CLA_BANCO</td>
					<td>CTA_BANCO</td><td>CLABE_INTERBANCARIA</td><td>RH_DET_POSICION_PLANTILLA</td><td>LUGAR_NAC</td><td>FECHA_NAC</td><td>SEXO</td>
					<td>SIND</td><td>CTA_CORREO</td><td>EDO_CIVIL</td><td>RFC</td><td>NUM_IMSS</td><td>CURP</td>
					<td>CALLE</td><td>COLONIA</td><td>CP</td><td>CIUDAD</td><td>ESTADO</td><td>CLA_DELEGACION</td>
					<td>TELEFONO</td><td>NACIONALIDAD</td><td>CLA_RAZON_SOCIAL</td><td>DIAS_CONT</td><td>CLA_TAB_SUE</td><td>NIV_TAB_SUE</td>
					<td>Global_ID</td><td>FECHA_SAL</td><td>FECHA_PER_INFO</td>
				</tr>";
			}
			for ($row = 2; $row <= $highestRow; $row++){
				$total = 0;
				$entro = 0;
				//A
				if($sheet->getCell(chr($column1).$row)==""){$total=1;}
				if($sheet->getCell(chr($column2).$row)==""){$total=1;}
				if($sheet->getCell(chr($column3).$row)==""){$total=1;}
				if($sheet->getCell(chr($column4).$row)==""){$total=1;}
				if($sheet->getCell(chr($column5).$row)==""){$total=1;}
				if($sheet->getCell(chr($column6).$row)==""){$total=1;}
				if($sheet->getCell(chr($column7).$row)==""){$total=1;}
				if($sheet->getCell(chr($column8).$row)==""){$total=1;}
				if($sheet->getCell(chr($column9).$row)==""){$total=1;}
				if($sheet->getCell(chr($column10).$row)==""){$total=1;}
				if($sheet->getCell(chr($column11).$row)==""){$total=1;}
				if($sheet->getCell(chr($column12).$row)==""){$total=1;}
				if($sheet->getCell(chr($column13).$row)==""){$total=1;}
				if($sheet->getCell(chr($column14).$row)==""){$total=1;}
				if($sheet->getCell(chr($column15).$row)==""){$total=1;}
				if($sheet->getCell(chr($column16).$row)==""){$total=1;}
				if($sheet->getCell(chr($column17).$row)==""){$total=1;}
				if($sheet->getCell(chr($column18).$row)==""){$total=1;}
				if($sheet->getCell(chr($column19).$row)==""){$total=1;}
				if($sheet->getCell(chr($column20).$row)==""){$total=1;}
				if($sheet->getCell(chr($column21).$row)==""){$total=1;}
				if($sheet->getCell(chr($column22).$row)==""){$total=1;}
				if($sheet->getCell(chr($column23).$row)==""){$total=1;}
				/*if($sheet->getCell(chr($column24).$row)==""){$total=1;}
				if($sheet->getCell(chr($column25).$row)==""){$total=1;}
				if($sheet->getCell(chr($column26).$row)==""){$total=1;}*/
				if($sheet->getCell(chr($column1).chr($column1).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column2).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column3).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column4).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column5).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column6).$row)==""){$total=1;}
				//if($sheet->getCell(chr($column1).chr($column7).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column8).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column9).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column10).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column11).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column12).$row)==""){$total=1;}
				/*if($sheet->getCell(chr($column1).chr($column13).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column14).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column15).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column16).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column17).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column18).$row)==""){$total=1;}*/
				if($sheet->getCell(chr($column1).chr($column19).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column20).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column21).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column22).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column23).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column24).$row)==""){$total=1;}
				if($sheet->getCell(chr($column1).chr($column25).$row)==""){$total=1;}
				/********************************************************************/
				/************************ VALIDACIONES ******************************/
				/********************************************************************/
				//Si el movimiento es un reingreso entra aqui
				//echo $sheet->getCell(chr($column1).$row)."</br>";
				if($sheet->getCell(chr($column1).$row)=="2" && $total==0){
					$sql = "SELECT count(*) as total 
					FROM rh_hist_laboral 
					WHERE CLA_TIPO_MOV=2
					AND CLA_TRAB=".$sheet->getCell(chr($column3).$row)."
					AND FECHA_MOV=cast('".$sheet->getCell(chr($column2).$row)."' AS DATETIME)";
					$stmt = sqlsrv_query($conn, $sql);
					$entro = 1;
				}else{
					//Si el movimiento es un PSMS entra aqui
					if($sheet->getCell(chr($column1).$row)=="3" && $total==0){
						$sql = "SELECT count(*) as total 
						FROM rh_hist_laboral 
						WHERE CLA_TIPO_MOV=3
						AND CLA_TRAB=".$sheet->getCell(chr($column3).$row)."
						AND FECHA_MOV=cast('".$sheet->getCell(chr($column2).$row)."' AS DATETIME)
						AND CLA_DEPTO=".$sheet->getCell(chr($column8).$row)."
						AND CLA_UBICACION=".$sheet->getCell(chr($column9).$row)."
						AND CLA_UBICACION_PAGO=".$sheet->getCell(chr($column10).$row)."
						AND CLA_CENTRO_COSTO=".$sheet->getCell(chr($column11).$row)."
						AND CLA_PUESTO=".$sheet->getCell(chr($column12).$row)."
						AND CLA_TAB_PRE=".$sheet->getCell(chr($column14).$row)."
						AND CLA_PERIODO=".$sheet->getCell(chr($column15).$row)."
						AND CLA_REG_IMSS=".$sheet->getCell(chr($column16).$row)."
						AND TIPO_CONTRATO=".$sheet->getCell(chr($column22).$row)."
						AND CLA_FORMA_PAGO=".$sheet->getCell(chr($column23).$row)."
						AND CLA_RAZON_SOCIAL=".$sheet->getCell(chr($column1).chr($column19).$row);
						$stmt = sqlsrv_query($conn, $sql);
						$entro = 1;
					}else{
						//Si el movimiento es un PCMS entra aqui
						if($sheet->getCell(chr($column1).$row)=="4" && $total==0){
							$sql = "SELECT count(*) as total 
							FROM rh_hist_laboral 
							WHERE CLA_TIPO_MOV=4
							AND CLA_TRAB=".$sheet->getCell(chr($column3).$row)."
							AND FECHA_MOV=cast('".$sheet->getCell(chr($column2).$row)."' AS DATETIME)
							AND CLA_DEPTO=".$sheet->getCell(chr($column8).$row)."
							AND CLA_UBICACION=".$sheet->getCell(chr($column9).$row)."
							AND CLA_UBICACION_PAGO=".$sheet->getCell(chr($column10).$row)."
							AND CLA_CENTRO_COSTO=".$sheet->getCell(chr($column11).$row)."
							AND CLA_PUESTO=".$sheet->getCell(chr($column12).$row)."
							AND CLA_TAB_PRE=".$sheet->getCell(chr($column14).$row)."
							AND CLA_PERIODO=".$sheet->getCell(chr($column15).$row)."
							AND CLA_REG_IMSS=".$sheet->getCell(chr($column16).$row)."
							AND TIPO_CONTRATO=".$sheet->getCell(chr($column22).$row)."
							AND CLA_FORMA_PAGO=".$sheet->getCell(chr($column23).$row)."
							AND CLA_RAZON_SOCIAL=".$sheet->getCell(chr($column1).chr($column19).$row)."
							AND SUELDO_MENSUAL>=".$sheet->getCell(chr($column19).$row);
							$stmt = sqlsrv_query($conn, $sql);
							$entro = 1;
						}
					}
				}
				//echo "Entro-->".$entro."</br>";
				if($entro==1){
					while( $row2 = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC) ) {
						$total = $row2['total'];
					}
				}
				/********************PROCESAMIENTO DE ARCHIVO DE EXCEL *********************************/
				//echo "Total-->".$total."</br>";
				if($total==0){
					$lineData= array(
						$sheet->getCell(chr($column1).$row),
						$sheet->getCell(chr($column2).$row),
						//CLA_TRAB
						$sheet->getCell(chr($column3).$row),
						//CLA_EMPRESA
						$sheet->getCell(chr($column4).$row),
						//AP_PATERNO
						$sheet->getCell(chr($column5).$row),
						//AP_MATERNO
						$sheet->getCell(chr($column6).$row),
						//NOM_TRAB
						$sheet->getCell(chr($column7).$row),
						//CLA_DEPTO
						$sheet->getCell(chr($column8).$row),
						//CLA_UBICACION
						$sheet->getCell(chr($column9).$row),
						//CLA_UBICACION_PAGO
						$sheet->getCell(chr($column10).$row),
						//CLA_CENTRO_COSTO
						$sheet->getCell(chr($column11).$row),
						//CLA_PUESTO
						$sheet->getCell(chr($column12).$row),
						//CLA_CLASIFICACION
						$sheet->getCell(chr($column13).$row),
						//CLA_TAB_PRE
						$sheet->getCell(chr($column14).$row),
						//CLA_PERIODO
						$sheet->getCell(chr($column15).$row),
						//CLA_REG_IMSS
						$sheet->getCell(chr($column16).$row),
						//TIPO_SALARIO
						$sheet->getCell(chr($column17).$row),
						//SUELDO_DIA
						$sheet->getCell(chr($column18).$row),
						//SUELDO_MENSUAL
						$sheet->getCell(chr($column19).$row),
						//FECHA_ING
						$sheet->getCell(chr($column20).$row),
						//FECHA_ING_GRUPO
						$sheet->getCell(chr($column21).$row),
						//TIPO_CONTRATO
						$sheet->getCell(chr($column22).$row),
						//CLA_FORMA_PAGO
						$sheet->getCell(chr($column23).$row),
						//CLA_BANCO
						$sheet->getCell(chr($column24).$row),
						//CTA_BANCO
						$sheet->getCell(chr($column25).$row),
						//CLABE_INTERBANCARIA
						$sheet->getCell(chr($column26).$row),
						//RH_DET_POSICION_PLANTILLA
						$sheet->getCell(chr($column1).chr($column1).$row),
						//LUGAR_NAC
						$sheet->getCell(chr($column1).chr($column2).$row),
						//FECHA_NAC
						$sheet->getCell(chr($column1).chr($column3).$row),
						//SEXO
						$sheet->getCell(chr($column1).chr($column4).$row),
						//SIND
						$sheet->getCell(chr($column1).chr($column5).$row),
						//CTA_CORREO
						$sheet->getCell(chr($column1).chr($column6).$row),
						//EDO_CIVIL
						$sheet->getCell(chr($column1).chr($column7).$row),
						//RFC
						$sheet->getCell(chr($column1).chr($column8).$row),
						//NUM_IMSS
						$sheet->getCell(chr($column1).chr($column9).$row),
						//CURP
						$sheet->getCell(chr($column1).chr($column10).$row),
						//CALLE
						$sheet->getCell(chr($column1).chr($column11).$row),
						//COLONIA
						$sheet->getCell(chr($column1).chr($column12).$row),
						//CP
						$sheet->getCell(chr($column1).chr($column13).$row),
						//CIUDAD
						$sheet->getCell(chr($column1).chr($column14).$row),															
						//ESTADO
						$sheet->getCell(chr($column1).chr($column15).$row),
						//CLA_DELEGACION
						$sheet->getCell(chr($column1).chr($column16).$row),
						//TELEFONO
						$sheet->getCell(chr($column1).chr($column17).$row),
						//NACIONALIDAD
						$sheet->getCell(chr($column1).chr($column18).$row),
						//CLA_RAZON_SOCIAL
						$sheet->getCell(chr($column1).chr($column19).$row),
						//DIAS_CON
						$sheet->getCell(chr($column1).chr($column20).$row),
						//CLA_TAB_SUE
						$sheet->getCell(chr($column1).chr($column21).$row),
						//NIV_TAB_SUE
						$sheet->getCell(chr($column1).chr($column22).$row),
						//Global_ID
						$sheet->getCell(chr($column1).chr($column23).$row),
						//FECHA_SAL
						$sheet->getCell(chr($column1).chr($column24).$row),
						//FECHA_PER_INFO
						$sheet->getCell(chr($column1).chr($column25).$row)
					);
					$tabla .= "<tr>";
						//CLA_TIPO_MOV
						$tabla.= "<td>".$sheet->getCell(chr($column1).$row)."</td>";
						//FECHA_MOV
						$tabla.= "<td>".$sheet->getCell(chr($column2).$row)."</td>";
						//CLA_TRAB
						$tabla.= "<td>".$sheet->getCell(chr($column3).$row)."</td>";
						//CLA_EMPRESA
						$tabla.= "<td>".$sheet->getCell(chr($column4).$row)."</td>";
						//AP_PATERNO
						$tabla.= "<td>".$sheet->getCell(chr($column5).$row)."</td>";
						//AP_MATERNO
						$tabla.= "<td>".$sheet->getCell(chr($column6).$row)."</td>";
						//NOM_TRAB
						$tabla.= "<td>".$sheet->getCell(chr($column7).$row)."</td>";
						//CLA_DEPTO
						$tabla.= "<td>".$sheet->getCell(chr($column8).$row)."</td>";
						//CLA_UBICACION
						$tabla.= "<td>".$sheet->getCell(chr($column9).$row)."</td>";
						//CLA_UBICACION_PAGO
						$tabla.= "<td>".$sheet->getCell(chr($column10).$row)."</td>";
						//CLA_CENTRO_COSTO
						$tabla.= "<td>".$sheet->getCell(chr($column11).$row)."</td>";
						//CLA_PUESTO
						$tabla.= "<td>".$sheet->getCell(chr($column12).$row)."</td>";
						//CLA_CLASIFICACION
						$tabla.= "<td>".$sheet->getCell(chr($column13).$row)."</td>";
						//CLA_TAB_PRE
						$tabla.= "<td>".$sheet->getCell(chr($column14).$row)."</td>";
						//CLA_PERIODO
						$tabla.= "<td>".$sheet->getCell(chr($column15).$row)."</td>";
						//CLA_REG_IMSS
						$tabla.= "<td>".$sheet->getCell(chr($column16).$row)."</td>";
						//TIPO_SALARIO
						$tabla.= "<td>".$sheet->getCell(chr($column17).$row)."</td>";
						//SUELDO_DIA
						$tabla.= "<td>".$sheet->getCell(chr($column18).$row)."</td>";
						//SUELDO_MENSUAL
						$tabla.= "<td>".$sheet->getCell(chr($column19).$row)."</td>";
						//FECHA_ING
						$tabla.= "<td>".$sheet->getCell(chr($column20).$row)."</td>";
						//FECHA_ING_GRUPO
						$tabla.= "<td>".$sheet->getCell(chr($column21).$row)."</td>";
						//TIPO_CONTRATO
						$tabla.= "<td>".$sheet->getCell(chr($column22).$row)."</td>";
						//CLA_FORMA_PAGO
						$tabla.= "<td>".$sheet->getCell(chr($column23).$row)."</td>";
						//CLA_BANCO
						$tabla.= "<td>".$sheet->getCell(chr($column24).$row)."</td>";
						//CTA_BANCO
						$tabla.= "<td>".$sheet->getCell(chr($column25).$row)."</td>";
						//CLABE_INTERBANCARIA
						$tabla.= "<td>".$sheet->getCell(chr($column26).$row)."</td>";
						//RH_DET_POSICION_PLANTILLA
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column1).$row)."</td>";
						//LUGAR_NAC
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column2).$row)."</td>";
						//FECHA_NAC
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column3).$row)."</td>";
						//SEXO
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column4).$row)."</td>";
						//SIND
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column5).$row)."</td>";
						//CTA_CORREO
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column6).$row)."</td>";
						//EDO_CIVIL
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column7).$row)."</td>";
						//RFC
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column8).$row)."</td>";
						//NUM_IMSS
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column9).$row)."</td>";
						//CURP
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column10).$row)."</td>";
						//CALLE
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column11).$row)."</td>";
						//COLONIA
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column12).$row)."</td>";
						//CP
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column13).$row)."</td>";
						//CIUDAD
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column14).$row)."</td>";															
						//ESTADO
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column15).$row)."</td>";
						//CLA_DELEGACION
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column16).$row)."</td>";
						//TELEFONO
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column17).$row)."</td>";
						//NACIONALIDAD
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column18).$row)."</td>";
						//CLA_RAZON_SOCIAL
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column19).$row)."</td>";
						//DIAS_CON
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column20).$row)."</td>";
						//CLA_TAB_SUE
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column21).$row)."</td>";
						//NIV_TAB_SUE
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column22).$row)."</td>";
						//Global_ID
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column23).$row)."</td>";
						//FECHA_SAL
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column24).$row)."</td>";
						//FECHA_PER_INFO
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column25).$row)."</td>";
					$tabla.= "</tr>";					
					fputcsv($f, $lineData, $delimiter);
				}
			}
		}
		//ftp_delete($connId, '/Entradalocal/'.$a);
		//$objFtp->delete($nombreArchivo);
		//Se elimina el archivo que se descargo del SFTP
		$strLocalFile = '/var/www/html/FTP/Descargas/'.$nombreArchivo;
		//unlink($strLocalFile);
		$contador++;
	}
}
//die();
//move back to beginning of file
fseek($f, 0);
$filename = "Archivo_" . date('Y-m-d') . ".csv";
//set headers to download file rather than displayed
header('Content-Type: text/csv');
header('Content-Disposition: attachment; filename="'.$filename . '";');
header('Content-Transfer-Encoding: binary');
header('Expires: 0');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
readfile('/Descargas/'.$filename);
//output all remaining data on a file pointer
fpassthru($f);
$tabla.= "</table>";
//echo $tabla;
// close the connection
//ftp_close($connId);
$objFtp->disconnect();
die();
?>