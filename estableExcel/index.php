<?PHP
header("Content-type: application/x-msdownload");
header("Content-Disposition: attachment; filename=Archivo_".date('Y-m-d').".xls");
require_once 'PHPExcel/Classes/PHPExcel.php';
// FTP server details
$ftpHost   = '192.168.56.1';
$ftpUsername = 'prueba';
$ftpPassword = 'asd123';

// open an FTP connection
$connId = ftp_connect($ftpHost) or die("Couldn't connect to $ftpHost");

// login to FTP server
$ftpLogin = ftp_login($connId, $ftpUsername, $ftpPassword);

$contents = ftp_mlsd($connId, "Entradalocal/");
/*var_dump($contents); // Aquí puedes recorrer el array devuelto para mostrar los ficheros.
die();*/
$contador=0;
$nombreArchivoArray = array();
$tabla="";

//SE COMPARA CONTRA LA BASE DE DATOS
$serverName = "10.11.0.34";
$connectionInfo = array( "Database"=>"Royal", "UID"=>"carlosj", "PWD"=>"martha01");
$conn = sqlsrv_connect( $serverName, $connectionInfo);

foreach($contents as $a){
	$nombreArchivo = $a['name'];
	$nombreArchivoArray[$contador] = $a['name'];
	$nombreTemp = explode(".", $nombreArchivo);
	$remoteFilePath = '/Entradalocal/'.$nombreArchivo;
	if(ftp_get($connId, "./Descargas/".$a['name'], $remoteFilePath, FTP_BINARY)){
		$localFilePath  = "/Descargas";
		//*******************************************************************************//
		//**************************SE PROCESA EL ARCHIVO********************************//
		//*******************************************************************************//
		$archivo = "./Descargas/".$nombreArchivo;
		$inputFileType = PHPExcel_IOFactory::identify($archivo);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
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
			if($sheet->getCell(chr($column24).$row)==""){$total=1;}
			if($sheet->getCell(chr($column25).$row)==""){$total=1;}
			if($sheet->getCell(chr($column26).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column1).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column2).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column3).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column4).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column5).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column6).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column7).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column8).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column9).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column10).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column11).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column12).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column13).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column14).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column15).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column16).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column17).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column18).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column19).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column20).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column21).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column22).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column23).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column24).$row)==""){$total=1;}
			if($sheet->getCell(chr($column1).chr($column25).$row)==""){$total=1;}
			//Si el movimiento es un reingreso entra aqui
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
			
			if($entro==1){
				while( $row2 = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC) ) {
					$total = $row2['total'];
				}
			}
			/********************PROCESAMIENTO DE ARCHIVO DE EXCEL *********************************/
			if($total==0){
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
			}
		}
	}
	ftp_delete($connId, '/Entradalocal/'.$a['name']);
	$contador++;
}
$tabla.= "</table>";
echo $tabla;
// close the connection
ftp_close($connId);
die();
?>