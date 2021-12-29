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
//$strPassword = 'Ayp$QuetOornOcdeb8';
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
}else{
	echo "Login success!!!/n";
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
$connectionInfo = array( "Database"=>"PPTEST", "UID"=>"carlosj", "PWD"=>"martha01");
$conn = sqlsrv_connect( $serverName, $connectionInfo);
//$arrayGeneral = array();
/*if($conn){
	echo "OK";
}else{
	echo "MAL";
}
die();*/
$localFilePath  = "/var/www/html/Descargas/";
$x = 0;
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
						//echo $sheet->getCell(chr($column1).$row)."</br>";
						//CLA_TIPO_MOV
						$arrayGeneral[$x]['CLA_TIPO_MOV'] = $sheet->getCell(chr($column1).$row)->getValue();
						//echo "Equis-->".$x."<-->".$arrayGeneral[$x]['CLA_TIPO_MOV']."</br>";
						$tabla.= "<td>".$sheet->getCell(chr($column1).$row)."</td>";
						//FECHA_MOV
						$arrayGeneral[$x]['FECHA_MOV'] = $sheet->getCell(chr($column2).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column2).$row)."</td>";
						//CLA_TRAB
						$arrayGeneral[$x]['CLA_TRAB'] = $sheet->getCell(chr($column3).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column3).$row)."</td>";
						//CLA_EMPRESA
						$arrayGeneral[$x]['CLA_EMPRESA'] = $sheet->getCell(chr($column4).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column4).$row)."</td>";
						//AP_PATERNO
						$arrayGeneral[$x]['AP_PATERNO'] = $sheet->getCell(chr($column5).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column5).$row)."</td>";
						//AP_MATERNO
						$arrayGeneral[$x]['AP_MATERNO'] = $sheet->getCell(chr($column6).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column6).$row)."</td>";
						//NOM_TRAB
						$arrayGeneral[$x]['NOM_TRAB'] = $sheet->getCell(chr($column7).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column7).$row)."</td>";
						//CLA_DEPTO
						$arrayGeneral[$x]['CLA_DEPTO'] = $sheet->getCell(chr($column8).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column8).$row)."</td>";
						//CLA_UBICACION
						$arrayGeneral[$x]['CLA_UBICACION'] = $sheet->getCell(chr($column9).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column9).$row)."</td>";
						//CLA_UBICACION_PAGO
						$arrayGeneral[$x]['CLA_UBICACION_PAGO'] = $sheet->getCell(chr($column10).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column10).$row)."</td>";
						//CLA_CENTRO_COSTO
						$arrayGeneral[$x]['CLA_CENTRO_COSTO'] = $sheet->getCell(chr($column11).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column11).$row)."</td>";
						//CLA_PUESTO
						$arrayGeneral[$x]['CLA_PUESTO'] = $sheet->getCell(chr($column12).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column12).$row)."</td>";
						//CLA_CLASIFICACION
						$arrayGeneral[$x]['CLA_CLASIFICACION'] = $sheet->getCell(chr($column13).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column13).$row)."</td>";
						//CLA_TAB_PRE
						$arrayGeneral[$x]['CLA_TAB_PRE'] = $sheet->getCell(chr($column14).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column14).$row)."</td>";
						//CLA_PERIODO
						$arrayGeneral[$x]['CLA_PERIODO'] = $sheet->getCell(chr($column15).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column15).$row)."</td>";
						//CLA_REG_IMSS
						$arrayGeneral[$x]['CLA_REG_IMSS'] = $sheet->getCell(chr($column16).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column16).$row)."</td>";
						//TIPO_SALARIO
						$arrayGeneral[$x]['TIPO_SALARIO'] = $sheet->getCell(chr($column17).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column17).$row)."</td>";
						//SUELDO_DIA
						$arrayGeneral[$x]['SUELDO_DIA'] = $sheet->getCell(chr($column18).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column18).$row)."</td>";
						//SUELDO_MENSUAL
						$arrayGeneral[$x]['SUELDO_MENSUAL'] = $sheet->getCell(chr($column19).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column19).$row)."</td>";
						//FECHA_ING
						$arrayGeneral[$x]['FECHA_ING'] = $sheet->getCell(chr($column20).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column20).$row)."</td>";
						//FECHA_ING_GRUPO
						$arrayGeneral[$x]['FECHA_ING_GRUPO'] = $sheet->getCell(chr($column21).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column21).$row)."</td>";
						//TIPO_CONTRATO
						$arrayGeneral[$x]['TIPO_CONTRATO'] = $sheet->getCell(chr($column22).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column22).$row)."</td>";
						//CLA_FORMA_PAGO
						$arrayGeneral[$x]['CLA_FORMA_PAGO'] = $sheet->getCell(chr($column23).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column23).$row)."</td>";
						//CLA_BANCO
						$arrayGeneral[$x]['CLA_BANCO'] = $sheet->getCell(chr($column24).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column24).$row)."</td>";
						//CTA_BANCO
						$arrayGeneral[$x]['CTA_BANCO'] = $sheet->getCell(chr($column25).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column25).$row)."</td>";
						//CLABE_INTERBANCARIA
						$arrayGeneral[$x]['CLABE_INTERBANCARIA'] = $sheet->getCell(chr($column26).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column26).$row)."</td>";
						//RH_DET_POSICION_PLANTILLA
						$arrayGeneral[$x]['RH_DET_POSICION_PLANTILLA'] = $sheet->getCell(chr($column1).chr($column1).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column1).$row)."</td>";
						//LUGAR_NAC
						$arrayGeneral[$x]['LUGAR_NAC'] = $sheet->getCell(chr($column1).chr($column2).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column2).$row)."</td>";
						//FECHA_NAC
						$arrayGeneral[$x]['FECHA_NAC'] = $sheet->getCell(chr($column1).chr($column3).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column3).$row)."</td>";
						//SEXO
						$arrayGeneral[$x]['SEXO'] = $sheet->getCell(chr($column1).chr($column4).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column4).$row)."</td>";
						//SIND
						$arrayGeneral[$x]['SIND'] = $sheet->getCell(chr($column1).chr($column5).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column5).$row)."</td>";
						//CTA_CORREO
						$arrayGeneral[$x]['CTA_CORREO'] = $sheet->getCell(chr($column1).chr($column6).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column6).$row)."</td>";
						//EDO_CIVIL
						$arrayGeneral[$x]['EDO_CIVIL'] = $sheet->getCell(chr($column1).chr($column7).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column7).$row)."</td>";
						//RFC
						$arrayGeneral[$x]['RFC'] = $sheet->getCell(chr($column1).chr($column8).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column8).$row)."</td>";
						//NUM_IMSS
						$arrayGeneral[$x]['NUM_IMSS'] = $sheet->getCell(chr($column1).chr($column9).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column9).$row)."</td>";
						//CURP
						$arrayGeneral[$x]['CURP'] = $sheet->getCell(chr($column1).chr($column10).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column10).$row)."</td>";
						//CALLE
						$arrayGeneral[$x]['CALLE'] = $sheet->getCell(chr($column1).chr($column11).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column11).$row)."</td>";
						//COLONIA
						$arrayGeneral[$x]['COLONIA'] = $sheet->getCell(chr($column1).chr($column12).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column12).$row)."</td>";
						//CP
						$arrayGeneral[$x]['CP'] = $sheet->getCell(chr($column1).chr($column13).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column13).$row)."</td>";
						//CIUDAD
						$arrayGeneral[$x]['CIUDAD'] = $sheet->getCell(chr($column1).chr($column14).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column14).$row)."</td>";															
						//ESTADO
						$arrayGeneral[$x]['ESTADO'] = $sheet->getCell(chr($column1).chr($column15).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column15).$row)."</td>";
						//CLA_DELEGACION
						$arrayGeneral[$x]['CLA_DELEGACION'] = $sheet->getCell(chr($column1).chr($column16).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column16).$row)."</td>";
						//TELEFONO
						$arrayGeneral[$x]['TELEFONO'] = $sheet->getCell(chr($column1).chr($column17).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column17).$row)."</td>";
						//NACIONALIDAD
						$arrayGeneral[$x]['NACIONALIDAD'] = $sheet->getCell(chr($column1).chr($column18).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column18).$row)."</td>";
						//CLA_RAZON_SOCIAL
						$arrayGeneral[$x]['CLA_RAZON_SOCIAL'] = $sheet->getCell(chr($column1).chr($column19).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column19).$row)."</td>";
						//DIAS_CON
						$arrayGeneral[$x]['DIAS_CON'] = $sheet->getCell(chr($column1).chr($column20).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column20).$row)."</td>";
						//CLA_TAB_SUE
						$arrayGeneral[$x]['CLA_TAB_SUE'] = $sheet->getCell(chr($column1).chr($column21).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column21).$row)."</td>";
						//NIV_TAB_SUE
						$arrayGeneral[$x]['NIV_TAB_SUE'] = $sheet->getCell(chr($column1).chr($column22).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column22).$row)."</td>";
						//Global_ID
						$arrayGeneral[$x]['Global_ID'] = $sheet->getCell(chr($column1).chr($column23).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column23).$row)."</td>";
						//FECHA_SAL
						$arrayGeneral[$x]['FECHA_SAL'] = $sheet->getCell(chr($column1).chr($column24).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column24).$row)."</td>";
						//FECHA_PER_INFO
						$arrayGeneral[$x]['FECHA_PER_INFO'] = $sheet->getCell(chr($column1).chr($column25).$row)->getValue();
						$tabla.= "<td>".$sheet->getCell(chr($column1).chr($column25).$row)."</td>";
					$tabla.= "</tr>";
					$x++;
					//fputcsv($f, $lineData, $delimiter);
				}
			}
		}
		//ftp_delete($connId, '/Entradalocal/'.$a);
		//$objFtp->delete($nombreArchivo);
		//Se elimina el archivo que se descargo del SFTP
		//$strLocalFile = '/var/www/html/FTP/Descargas/'.$nombreArchivo;
		//unlink($strLocalFile);
		$contador++;
	}
}
echo "<pre>";
print_r($arrayGeneral);
echo "</pre></br></br>";
//die();*/

$intArray[0]['entero_default'] = 0;
$intArray[0]['float_default'] = 0.00;
$intArray[0]['varchar'] = '';

$contador=0;
$serverName2 = "10.11.0.34"; //serverName\instanceName
//$connectionInfo2 = array( "Database"=>"Royal", "UID"=>"", "PWD"=>"");
$connectionInfo2 = array( "Database"=>"PPTEST", "UID"=>"carlosj", "PWD"=>"martha01");
$conn2 = sqlsrv_connect( $serverName2, $connectionInfo2);

/*$sql2 = "{call sp_KER_AltaTrabajadores_WinServ
	(
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
		?, ?, ?, ?, ?
	)}";*/

foreach($arrayGeneral as $a){
	$sql2 = "exec sp_KER_AltaTrabajadores_WinServ 
	@ClaTrab = ".$a['CLA_TRAB'].", 
	@ClaEmpresa = ".$a['CLA_EMPRESA'].", 
	@ApPaterno = '".$a['AP_PATERNO']."',
	@ApMaterno = '".$a['AP_MATERNO']."', 
	@Nombre = '".$a['NOM_TRAB']."', 
	@ClaUbBase = ".$a['CLA_UBICACION'].",  
	@ClaUbPago = ".$a['CLA_UBICACION_PAGO'].",
	@ClaDepto = ".$a['CLA_DEPTO'].",  
	@ClaPuesto = ".$a['CLA_PUESTO'].",  
	@ClaClasif = ".$a['CLA_CLASIFICACION'].",  
	@ClaRolTurno = ".$intArray[0]['entero_default'].",  
	@ClaTabPre = ".$a['CLA_TAB_PRE'].",  
	@ClaPeriodo = ".$a['CLA_PERIODO'].", 
	@ClaRegPat = '".$a['CLA_REG_IMSS']."',  
	@ClaCC = ".$a['CLA_CENTRO_COSTO'].",
	@TipoSal = ".$a['TIPO_SALARIO'].",  
	@SueDiario = '".$a['SUELDO_DIA']."',  
	@SueInteg  = '".$intArray[0]['float_default']."', 
	@SueMensual = '".$a['SUELDO_MENSUAL']."',  
	@FechaIngreso = '".$a['FECHA_ING']."', 
	@FechaIngresoEmpresa = '".$a['FECHA_ING_GRUPO']."', 
	@TipoContrato = ".$a['TIPO_CONTRATO'].", 
	@ClaFormaPago = ".$a['CLA_FORMA_PAGO'].",  
	@ClaBanco = '".$a['CLA_BANCO']."',  
	@CuentaBanco = '".$a['CTA_BANCO']."', 
	@LugarNac = '".$a['LUGAR_NAC']."', 
	@FechaNac = '".$a['FECHA_NAC']."',  
	@Sexo = '".$a['SEXO']."', 
	@Sindicalizado = ".$a['SIND'].", 
	@Email = '".$a['CTA_CORREO']."', 
	@EdoCivil = '".$a['EDO_CIVIL']."', 
	@RFC = '".$a['RFC']."', 
	@NoIMSS = '".$a['NUM_IMSS']."', 
	@CURP = '".$a['CURP']."', 
	@Calle = '".$a['CALLE']."',  
	@Colonia = '".$a['COLONIA']."', 
	@CodPostal = ".$a['CP'].",
	@Ciudad = '".$a['CIUDAD']."', 
	@Estado = '".$a['ESTADO']."', 
	@Telefono = '".$a['TELEFONO']."',	
	@EstatusDelTrabajador = ".$intArray[0]['entero_default'].", 
	@FechaDeBaja = '".$intArray[0]['varchar']."', 
	@NumeroDeCartillaMilitar = '".$intArray[0]['varchar']."', 
	@Nacionalidad = '".$intArray[0]['varchar']."', 
	@Estatura = '".$intArray[0]['float_default']."', 
	@Peso = '".$intArray[0]['float_default']."', 
	@NombredelPadre = '".$intArray[0]['varchar']."', 
	@NombredelaMadre = '".$intArray[0]['varchar']."', 
	@AvisarenCasodeAccidente = '".$intArray[0]['varchar']."', 
	@AvisoParentesco = ".$intArray[0]['entero_default'].", 
	@AvisoCalle = '".$intArray[0]['varchar']."',  
	@AvisoCiudad = '".$intArray[0]['varchar']."',  
	@AvisoCP = ".$intArray[0]['entero_default'].",  
	@AvisoTelefono = '".$intArray[0]['varchar']."',	
	@ClaveRazonsocial = ".$a['CLA_RAZON_SOCIAL'].",
	@Plaza = ".$intArray[0]['entero_default'].", 
	@FinContrato = '".$intArray[0]['varchar']."',  
	@ClaTipoMov = ".$a['CLA_TIPO_MOV'].", 
	@EstadoNacimiento = '".$intArray[0]['varchar']."',  
	@Familiares			= '".$intArray[0]['varchar']."',  
	@Escolaridad		= '".$intArray[0]['varchar']."',  
	@NumInfonavit		= '".$intArray[0]['varchar']."',  
	@pnGlobalID			= ".$a['Global_ID'].", 
	@pdtFechaCambSal	= '".$a['FECHA_SAL']."', 
	@pdtFechaCambEstru  = '".$intArray[0]['varchar']."'";
	echo $sql2."</br>GO</br>";
	//die();
	/*if($contador==7){
		echo "<pre>".$sql2."</pre>"."</br>";
		$stmt = sqlsrv_prepare( $conn2, $sql2, array());
		if( !$stmt ) {
			die( print_r( sqlsrv_errors(), true));
		}else{
			echo "prepare correct</br>";
		}

		$ejecucion = sqlsrv_execute($stmt);
		if( !$ejecucion ) {
			die(print_r(sqlsrv_errors(), true));
		}else{
			echo "OK";
			die();
		}
		die();
	}*/
	$contador++;
}
die();
?>