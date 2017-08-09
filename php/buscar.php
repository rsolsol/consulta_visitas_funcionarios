<?php
//coneccion al servidor para realizar la consulta
    $host = 'xxxxx';
    $user = 'xxxxx';
    $pass = 'xxxxx';
    $bd   = 'xxxxx';

    $coneccion = mysqli_connect($host, $user, $pass, $bd) or die ('no se Puede Conectar: '. mysqli_errno());
// fin de la conecion al servidor
require_once 'Classes/PHPExcel.php';

$fechaI = $_GET['fecha1'];
$fechaF = $_GET['fecha2'];

// cambia el formato de fecha para realizar la lectura en mysql//
//fecha de Inicio de Busqueda
$i=explode('/', $fechaI);
$fecha_sql_I=$i[2]."-".$i[1]."-".$i[0];

//fecha Final de Busqueda
$f=explode('/', $fechaF);
$fecha_sql_F=$f[2]."-".$f[1]."-".$f[0];

$script_fecha= "select fecha, hora_entrada as entrada, concat(apellidos,' ',nombre) as NombreCompleto, td_detalle_dominio.co_detalle_dominio as tipoDocumento, visita.nro_documento as  umeroDocumento, visitante.entidad as entidad, motivo, empleado.no_completo as empleado, empleado.cargo as cargo, hora_salida as salida, area.descripcion as oficina from tm_visitante visitante, tm_visitas visita, td_detalle_dominio, tm_empleado_publico empleado, tm_area area where cast(fecha AS date) BETWEEN '$fecha_sql_I' AND '$fecha_sql_F' and visita.nro_documento=visitante.Nro_documento and visita.cod_empleado=empleado.cod_empleado and td_detalle_dominio.id_detalle_dominio=visita.id_detalle_dominio and area.cod_area=empleado.cod_area and hora_salida is null order by fecha asc;";

$consulta = mysqli_query($coneccion, $script_fecha) or die ('Consulta fallida');
/*verifica que este saliendo la informacion 
while ($fila = mysqli_fetch_array($consulta)) {
          print_r($fila['NombreCompleto']."<br>");
	}
*/
date_default_timezone_set('America/Lima');

// Crea un nuevo objeto PHPExcel
    $objPHPExcel = new PHPExcel();
// Establecer propiedades
$objPHPExcel->getProperties()
    ->setCreator("Cattivo") // Nombre del autor
    ->setLastModifiedBy("Cattivo") //Ultimo usuario que lo modificó
    ->setTitle("Documento Excel de Prueba") // Titulo
    ->setSubject("Documento Excel de Prueba") //Asunto
    ->setDescription("Demostracion sobre como crear archivos de Excel desde PHP.") //Descripción
    ->setKeywords("Excel Office 2007 openxml php") //Etiquetas
    ->setCategory("Pruebas de Excel"); //Categorias

// Se agregan los datos de los alumnos    
    $i = 6;
    while ($fila = mysqli_fetch_array($consulta)) {
              
        $a=explode('-', $fila['fecha']);
        $fecha__I=$a[2]."-".$a[1]."-".$a[0];//se cambia al formato peruano d-m-y
        
        $objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('B'.$i, $fecha__I)
            ->setCellValue('C'.$i, $fila['entrada'])
            ->setCellValue('D'.$i, $fila['NombreCompleto'])
            ->setCellValue('E'.$i, $fila['umeroDocumento'])
            ->setCellValue('F'.$i, $fila['entidad'])
            ->setCellValue('G'.$i, $fila['motivo'])
            ->setCellValue('H'.$i, $fila['empleado'])
            ->setCellValue('I'.$i, $fila['cargo'])
            ->setCellValue('J'.$i, $fila['salida'])
            ->setCellValue('K'.$i, $fila['oficina']);
            $i++;
    }

// Agregar Informacion
    $entidad = "Municipalidad Distrital de Puente Piedra";
    $tituloReporte = "Visitas a Funcionarios";
    $tituloPrincipal = "Reporte de Visitas a Funcionarios del ".$fechaI." al ".$fechaF;
    $titulosColumnas = array('FECHA', 'ENTRADA', 'DATOS', 'DNI', 'ENTIDAD', 'MOTIVO', 'EMPLEADO', 'CARGO', 'SALIDA', 'OFICINA');
    $hoy = date('l jS \of F Y h:i:s A');
    $piedepagina = "Creado el ".$hoy;
// Agrega formato de estilo a los resultados
    $estiloEntidad = array(
        'font' => array(
            'bold'  => true, //coloca negrita
            'size' =>16,
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'rotation' => 0,
            'wrap' => TRUE
        )
    );

    $estiloReporte = array (
        'font' => array(
            'bold'  => true, //coloca negrita
            'size' =>16,
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'rotation' => 0,
            'wrap' => TRUE
        )
    );

    $estiloTituloPrincipal = array(
        'font' => array(
        'name'      => 'Verdana',
        'bold'      => true,
        'italic'    => false,
        'strike'    => false,
        'size' =>16,
        'color'     => array(
            'rgb' => 'FFFFFF'
            )   
        ),
        'fill' => array(
          'type'  => PHPExcel_Style_Fill::FILL_SOLID,
          'color' => array(
                'argb' => 'FF220835')
        ),
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_NONE
            )
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'rotation' => 0,
            'wrap' => TRUE
        )
    );

    $estiloTituloColumnas = array(
        'font' => array(
            'name'  => 'Arial',
            'bold'  => true,
            'color' => array(
                'rgb' => 'FFFFFF'
            )
        ),
        'fill' => array(
            'type'       => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
            'rotation'   => 90,
            'startcolor' => array(
                'rgb' => 'FF8C00'
                ),
            'endcolor' => array(
                'rgb' => 'ff6e40'
                )
        ),
        'borders' => array(
            'top' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
                'color' => array(
                    'rgb' => '143860'
                )
            ),
            'bottom' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM ,
                'color' => array(
                    'rgb' => '143860'
                )
            )
        ),
        'alignment' =>  array(
            'horizontal'=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical'  => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            'wrap'      => TRUE
        )
    );

    $estiloInformacion = new PHPExcel_Style();
    $estiloInformacion->applyFromArray( array(
        'font' => array(
            'name'  => 'Arial',
            'color' => array(
                'rgb' => '000000'
            )
        ),
        'fill' => array(
      'type'  => PHPExcel_Style_Fill::FILL_SOLID,
      'color' => array(
                'rgb' => 'fff8e1') //color de fondo de los items
      ),
        'borders' => array(
            'left' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN ,
          'color' => array(
                  'rgb' => '3a2a47'
                )
            )
        )
    ));

// Se combinan las celdas A1 hasta D1, para colocar ahí el titulo del reporte
    $objPHPExcel->setActiveSheetIndex(0)
    ->mergeCells('A1:D1');
    $objPHPExcel->setActiveSheetIndex(0)
    ->mergeCells('A3:K3');
    $objPHPExcel->setActiveSheetIndex(0)
    ->mergeCells('I1:K1');

//Se aplica los estilos dados
    $objPHPExcel->getActiveSheet()->getStyle('A1:D1')->applyFromArray($estiloEntidad);
    $objPHPExcel->getActiveSheet()->getStyle('I1:K1')->applyFromArray($estiloReporte);
    $objPHPExcel->getActiveSheet()->getStyle('A3:K3')->applyFromArray($estiloTituloPrincipal);
    $objPHPExcel->getActiveSheet()->getStyle('B5:K5')->applyFromArray($estiloTituloColumnas);
    $objPHPExcel->getActiveSheet()->setSharedStyle($estiloInformacion, "B6:K".($i-1));

// Se agregan los titulos del reporte
    $i++;
    $objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', $entidad)
    ->setCellValue('I1', $tituloReporte)
    ->setCellValue('A3', $tituloPrincipal)
    ->setCellValue('B5', $titulosColumnas[0])
    ->setCellValue('C5', $titulosColumnas[1])
    ->setCellValue('D5', $titulosColumnas[2])
    ->setCellValue('E5', $titulosColumnas[3])
    ->setCellValue('F5', $titulosColumnas[4])
    ->setCellValue('G5', $titulosColumnas[5])
    ->setCellValue('H5', $titulosColumnas[6])
    ->setCellValue('I5', $titulosColumnas[7])
    ->setCellValue('J5', $titulosColumnas[8])
    ->setCellValue('K5', $titulosColumnas[9])
    ->setCellValue('A'.$i, $piedepagina);
    /*->setCellValue('K'.$i, $fila['oficina']);*/

// Asignar el ancho de las columnas de forma automatica en base al contenido
    for($i = 'B'; $i <= 'K'; $i++){
        $objPHPExcel->setActiveSheetIndex(0)->getColumnDimension($i)->setAutoSize(TRUE);
    }
// Renombrar Hoja
    $objPHPExcel->getActiveSheet()->setTitle('Visitas-MDPP');

// Establecer la hoja activa, para que cuando se abra el documento se muestre primero.
    $objPHPExcel->setActiveSheetIndex(0);

// Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Visitas-MDPP.xlsx"');
    header('Cache-Control: max-age=0');
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('php://output');
    exit;
//http://comunidad.fware.pro/dev/php/como-crear-verdaderos-archivos-de-excel-usando-phpexcel/
//http://www.codedrinks.com/crear-un-reporte-en-excel-con-php-y-mysql/
?>