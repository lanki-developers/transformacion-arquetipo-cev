<?php
if($PCO_Accion=="Cargar_Excel")
{
    //echo "Hola 1";
    //Determina el tipo de archivo detectado
    $XLFileType = PHPExcel_IOFactory::identify($archivo_cargado);

    //Crear la instancia para leer el excel
    $leer_excel = PHPExcel_IOFactory::createReaderForFile($archivo_cargado);

    //Cargar Excel
    $excel_obj = $leer_excel->load($archivo_cargado);

    //Cargar hoja de trabajo
    $hoja_arquetipo_inicial = $excel_obj->getSheet(0);

    //obtener filas
    $filas = $hoja_arquetipo_inicial->getHighestRow();

    echo '<script>
      function actualizarFilaRevisada(fila){
        $("#fila_revisando").html("Revisadas "+fila+" ");
      }
    </script>';
    echo '<div><span id="fila_revisando"></span>de '.$filas.' filas<hr/>';
    //die();
    $meses = [
      'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
    ];

    $idx = 1;
    $evento_anterior = null;
    $contador_cursos = 0;
    $fila_html = "";

    for($row=2; $row<=$filas; $row++)
    {
      //Capturar valor de las columnas

      //echo "<br/><hr/>Filas $row";
      ?>
      <script>actualizarFilaRevisada(<?php echo $row; ?>)</script>
      <?php
      //$fecha_clase = transformarFechaClase($hoja_arquetipo_inicial->getCell('H'.$row)->getValue()); //COL-CURSO H
      //var_dump($fecha_clase);

      $clases_en_vivo[$row-1]['linea'] = $row; // Contar ls lineas para el mensaje de error
      $clases_en_vivo[$row-1]['codigocurso'] = $hoja_arquetipo_inicial->getCell('G'.$row)->getValue(); //COL-EVENTO G

      $clases_en_vivo[$row-1]['nombrecurso'] = $hoja_arquetipo_inicial->getCell('I'.$row)->getValue(); //COL-PROGRAMAS-Fechas I

      $clases_en_vivo[$row-1]['descripcion'] = transformacionDescripcionCurso($clases_en_vivo[$row-1]['nombrecurso']);

      $url_imagen_curso = transformacionUrlImagenCurso($clases_en_vivo[$row-1]['nombrecurso']);

      // echo $clases_en_vivo[$row-1]['codigocurso']."<br/>" ;
      // echo $clases_en_vivo[$row-1]['nombrecurso']."<br/>" ;
      // echo $clases_en_vivo[$row-1]['descripcion']."<br/>" ;

      //var_dump($url_imagen_curso); // url or false
      if( $url_imagen_curso )
        $clases_en_vivo[$row-1]['urlimagen'] = $url_imagen_curso;
      else
        $clases_en_vivo[$row-1]['urlimagen'] = null;

      //echo $clases_en_vivo[$row-1]['urlimagen']."<br/>" ;

      //Ajustar fecha de inicio
      $fecha_inicio = transformarFecha($hoja_arquetipo_inicial->getCell('K'.$row)); //COL-Fecha-Inicio K
      $clases_en_vivo[$row-1]['fecha_inicio'] = date('Y-m-d', PHPExcel_Shared_Date::ExcelToPHP($hoja_arquetipo_inicial->getCell('K'.$row)->getValue()));
      //echo "fecha_inicio:" . $clases_en_vivo[$row-1]['fecha_inicio'] . "<br/>";
      $clases_en_vivo[$row-1]['dia'] = $fecha_inicio['dia'];
      $clases_en_vivo[$row-1]['mes'] = $meses[$fecha_inicio['mes']-1];
      $clases_en_vivo[$row-1]['anio'] = $fecha_inicio['ano'];

      $clases_en_vivo[$row-1]['horario'] = $hoja_arquetipo_inicial->getCell('J'.$row)->getValue(); //COL-Horario J
      //echo "Horario" . $clases_en_vivo[$row-1]['horario'] . "<br/>";
      $horarios = transformarHorario($clases_en_vivo[$row-1]['fecha_inicio'],$hoja_arquetipo_inicial->getCell('J'.$row)->getValue());
      $clases_en_vivo[$row-1]['horai'] = $horarios['inicio']['hora'];
      $clases_en_vivo[$row-1]['minutoi'] = $horarios['inicio']['min'];
      $clases_en_vivo[$row-1]['horaf'] = $horarios['final']['hora'];
      $clases_en_vivo[$row-1]['minutof'] = $horarios['final']['min'];

      // echo $clases_en_vivo[$row-1]['fecha_inicio']."<br/>";
      // echo $clases_en_vivo[$row-1]['dia']. "/" . $clases_en_vivo[$row-1]['mes']. "/" . $clases_en_vivo[$row-1]['anio']."<br/>";
      // echo $clases_en_vivo[$row-1]['horai'].":" . $clases_en_vivo[$row-1]['minutoi']."<br/>";
      // echo $clases_en_vivo[$row-1]['horaf'].":" . $clases_en_vivo[$row-1]['minutof']."<br/>";

      $url_sala_webex = $hoja_arquetipo_inicial->getCell('V'.$row)->getValue(); //COL-sala webex V
      //var_dump($url_sala_webex);
      if( ! is_null($url_sala_webex)){
        $clases_en_vivo[$row-1]['webex'] = $url_sala_webex;

        $clases_en_vivo[$row-1]['sala'] = $hoja_arquetipo_inicial->getCell('W'.$row)->getValue(); //COL-sala y Contraseña W
        //$clases_en_vivo[$row-1]['contrasena'] = $hoja_arquetipo_inicial->getCell('V'.$row)->getValue();
        //$$clases_en_vivo[$row-1]['contrasena'] =

        //$html_code = explode(":", $clases_en_vivo[$row-1]['contrasena']);
        //$explode_contrasena = explode("Contraseña",$html_code[1]);
        $clases_en_vivo[$row-1]['salaycont'] = "<p>Número de la reunión (código de acceso): <b>".$clases_en_vivo[$row-1]['sala']."</b></p><p>Contraseña de la reunión: <b>".$clases_en_vivo[$row-1]['contrasena']."</b><br></p>";

      }else{
        $clases_en_vivo[$row-1]['webex'] = null;
        $clases_en_vivo[$row-1]['salaycont'] = null;
      }

      // echo $clases_en_vivo[$row-1]['webex']."<br/>" ;
      // echo $clases_en_vivo[$row-1]['salaycont']."<br/>";
      //
      // echo "evento_anterior:".$evento_anterior."<br/>";
      // echo "codigocurso:".$clases_en_vivo[$row-1]['codigocurso']."<br/>";

      if( is_null($evento_anterior) || $evento_anterior != $clases_en_vivo[$row-1]['codigocurso']){
        $idx_clase=1;
        //Agregar  a la lista de cursos
        $listado[$contador_cursos]['codigocurso'] = $clases_en_vivo[$row-1]['codigocurso'];
        $listado[$contador_cursos]['urlimagen'] = $clases_en_vivo[$row-1]['urlimagen'];
        $listado[$contador_cursos]['salaycont'] = $clases_en_vivo[$row-1]['salaycont'];
        $listado[$contador_cursos]['descripcion'] = $clases_en_vivo[$row-1]['descripcion'];
        //contar curso
        $contador_cursos++;
      }else{
        $idx_clase++;
      }
      $evento_anterior = $clases_en_vivo[$row-1]['codigocurso'];
      $clases_en_vivo[$row-1]['numero'] = $idx_clase;
      $clases_en_vivo[$row-1]['clase'] = 'Clase '.$idx_clase;

      //echo "evento_anterior_nuevo:".$evento_anterior."<br/>";
      //echo $clases_en_vivo[$row-1]['numero']."<br/>" ;
      //echo $clases_en_vivo[$row-1]['clase']."<br/>" ;

      $fecha_entrega = $hoja_arquetipo_inicial->getCell('A'.$row)->getValue();
      $interactividad = $hoja_arquetipo_inicial->getCell('D'.$row)->getValue();
      //echo "fecha_entrega:".$fecha_entrega."<br/>";
      //echo "interactividad:".$interactividad."<br/>";

      if( ! empty($fecha_entrega) && $interactividad=='Curso en vivo' )
      {
        $fila =  "<tr>" .
                  "<td>".$idx."</td>".
                  "<td>".$clases_en_vivo[$row-1]['codigocurso']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['numero']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['clase']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['dia']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['mes']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['anio']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['horai']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['minutoi']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['horaf']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['minutof']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['descripcion']."</td>".
                  "<td>".$clases_en_vivo[$row-1]['webex']."</td>".
              "</tr>";

              $fila_html .= $fila;
          //echo $fila;

        $idx++;
      }

      //echo $row . " de " . $filas . "->" . $interactividad . "->" . $evento_anterior . " => " . $idx . "<br/>";

    }
    //echo "<hr/>fin for <br/>";

    $rpta_listados_cursos = agregarRegistroTablaListadosCursos($listado);
    $rpta_clases_cursos = agregarRegistroTablaClasesCursosEnVivo($clases_en_vivo);

    echo '<div class="panel panel-primary">
            <div class="panel-heading">Importación Arquetipo</div>
            <div class="panel-body">
              <div class="alert alert-info" role="alert">
                <strong>Importación Realizada!</strong>'.
                $rpta_listados_cursos.$rpta_clases_cursos.'
              </div>' .
              '<div class="alert">
                <a role="button" class="btn btn-success"  href="index.php?PCO_Accion=CrearExcelArquetipoCursosEnVivo">
                <i class="fa fa-file-excel-o" aria-hidden="true"></i> Generar arquetipo</a>'.
              '</div>'.
              '<div class="alert alert-warning alert-dismissible" role="alert">
                <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <strong>Advertencia! </strong>Se han encontrado las siguientes inconsistencias y no fueron importadas:<br/><br/>'.
                encontrarErrores($clases_en_vivo).
              '</div>
            </div>'.
          '</div>';

}

if($PCO_Accion=="CrearExcelArquetipoCursosEnVivo")
{
  //echo "generar excel <br/>";
  $listados_cursos = NERA_TraerContenidoTabla("app_listados_cursos");
  //print_r($listados_cursos);

  $clases_cursos = NERA_TraerContenidoTabla("app_clases_cursos");
  //print_r($clases_cursos);

  $objPHPExcel = new PHPExcel();
  //Activar Hoja Uno
  $objPHPExcel->setActiveSheetIndex(0);
  $objPHPExcel->getActiveSheet()->setTitle("Cursos en Vivo");
  $rowCount = 1;
  //Incluir cabeceras
  $objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount, 'item');
  $objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount, 'codigocurso');
  $objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount, 'numero');
  $objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount, 'clase');
  $objPHPExcel->getActiveSheet()->SetCellValue('E'.$rowCount, 'dia');
  $objPHPExcel->getActiveSheet()->SetCellValue('F'.$rowCount, 'mes');
  $objPHPExcel->getActiveSheet()->SetCellValue('G'.$rowCount, 'ano');
  $objPHPExcel->getActiveSheet()->SetCellValue('H'.$rowCount, 'horai');
  $objPHPExcel->getActiveSheet()->SetCellValue('I'.$rowCount, 'minutoi');
  $objPHPExcel->getActiveSheet()->SetCellValue('J'.$rowCount, 'horaf');
  $objPHPExcel->getActiveSheet()->SetCellValue('K'.$rowCount, 'minutof');
  $objPHPExcel->getActiveSheet()->SetCellValue('L'.$rowCount, 'descripcion');
  $objPHPExcel->getActiveSheet()->SetCellValue('M'.$rowCount, 'webex');
  $rowCount++;

  foreach($clases_cursos as $row){
    $objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount, $rowCount-1);
    $objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount, $row['codigocurso']);
    $objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount, $row['numero']);
    $objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount, $row['clase']);
    $objPHPExcel->getActiveSheet()->SetCellValue('E'.$rowCount, $row['dia']);
    $objPHPExcel->getActiveSheet()->SetCellValue('F'.$rowCount, $row['mes']);
    $objPHPExcel->getActiveSheet()->SetCellValue('G'.$rowCount, $row['anio']);
    $objPHPExcel->getActiveSheet()->SetCellValue('H'.$rowCount, $row['horai']);
    $objPHPExcel->getActiveSheet()->SetCellValue('I'.$rowCount, $row['minutoi']);
    $objPHPExcel->getActiveSheet()->SetCellValue('J'.$rowCount, $row['horaf']);
    $objPHPExcel->getActiveSheet()->SetCellValue('K'.$rowCount, $row['minutof']);
    $objPHPExcel->getActiveSheet()->SetCellValue('L'.$rowCount, $row['descripcion']);
    $objPHPExcel->getActiveSheet()->SetCellValue('M'.$rowCount, $row['webex']);
    $rowCount++;
  }

  //Activar Hoja Dos
  $objPHPExcel->createSheet(NULL, 1);
  $objPHPExcel->setActiveSheetIndex(1);
  $objPHPExcel->getActiveSheet()->setTitle("Listado de Cursos");
  $rowCount = 1;
  //Incluir cabeceras
  $objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount, 'codigocurso2');
  $objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount, 'urlimagen');
  $objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount, 'salaycont');
  $objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount, 'descripcion');
  $rowCount++;
  foreach($listados_cursos as $row){
      $objPHPExcel->getActiveSheet()->SetCellValue('A'.$rowCount, $row['codigocurso']);
      $objPHPExcel->getActiveSheet()->SetCellValue('B'.$rowCount, $row['urlimagen']);
      $objPHPExcel->getActiveSheet()->SetCellValue('C'.$rowCount, $row['salaycont']);
      $objPHPExcel->getActiveSheet()->SetCellValue('D'.$rowCount, $row['descripcion']);
      $rowCount++;
  }

  $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
  $file_name_xlsx = './mod/importar_arquetipo/files/arquetipo_cursos_en_vivo_'.date("Ymd-H_i_s", time()).'.xlsx';
  $objWriter->save($file_name_xlsx);

  echo '<div class="panel panel-primary">
            <div class="panel-heading">Importación Arquetipo</div>
            <div class="panel-body">
              <div class="alert alert-info" role="alert">
                <strong>Descargar archivo</strong> '.
                'Se ha generado un archivo de arquetipo de cursos en vivo, para utilizar con el Robot.<br/>Haga clic en el botón para descargar.'.
              '</div>' .
              '<div class="alert">
                <a href="'.$file_name_xlsx.'" role="button" class="btn btn-info" style="margin-right: 5px;">
                <i class="fa fa-cloud-download" aria-hidden="true"></i> Descargar archivo</a>'.
               '<a href="index.php?PCO_Accion=PCO_CargarObjeto&PCO_Objeto=frm:6:1" role="button" class="btn btn-warning" style="margin-right: 5px;">
                <i class="fa fa-undo" aria-hidden="true"></i> Importar otro archivo</a>'.
              '</div>'.
            '</div>'.
        '</div>';
}

function transformarFechaClase($cell)
{
  //$InvDate = $cell->getValue()
  //$I
  $pos = strpos($mystring, $findme);

  // Nótese el uso de ===. Puesto que == simple no funcionará como se espera
  // porque la posición de 'a' está en el 1° (primer) caracter.
  if ($pos === false)
    return $pos ;

  $partes = explode("-",$cell);
  $fecha = trim($partes[1]);

  //$time = strtotime($fecha_clase);
  $fecha = DateTime::createFromFormat("d/m/Y", $fecha);
  //$fecha_clase = date('Y-m-d',$time);
  $fecha_clase['completa'] = $fecha->format('Y-m-d');
  $fecha_clase['dia'] = $fecha->format('j');
  $fecha_clase['mes'] = $fecha->format('n');
  $fecha_clase['ano'] = $fecha->format('Y');

  return $fecha_clase;
}

function transformarFecha($cell)
{
  $InvDate = $cell->getValue();
  if(PHPExcel_Shared_Date::isDateTime($cell)) {

    //$clases_en_vivo[$row-1]['fecha_inicio'] = date('Y-m-d', PHPExcel_Shared_Date::ExcelToPHP($InvDate));

    $fecha['dia'] = date('j',PHPExcel_Shared_Date::ExcelToPHP($InvDate));
    $fecha['mes'] = date('n',PHPExcel_Shared_Date::ExcelToPHP($InvDate));
    $fecha['ano'] = date('Y',PHPExcel_Shared_Date::ExcelToPHP($InvDate));

    return $fecha;
  }
  return false;
}

function transformacionUrlImagenCurso($value)
{
  $url_partes = explode("- 528",$value);
  $sql = "SELECT imagen_curso_url FROM app_imagenes_curso WHERE nombre_curso = '" .trim($url_partes[0])."' AND creada='Si'";
  $url_imagen = PCO_EjecutarSQL($sql)->fetchColumn();
  return $url_imagen;
}

function transformarHorario($fecha_clase,$value)
{
  $horario_partes = explode(" ",$value);
  $horas_partes = explode("-",$horario_partes[1]);
  $horario['inicio'] = explode(":",$horas_partes[0]);
  $horario['final'] = explode(":",$horas_partes[1]);
  //crear fecha y hora de inicio
  $fecha_inicio = $fecha_clase." ". $horario['inicio'][0] . ":". $horario['inicio'][1] . ":00";
  $fecha_inicio = DateTime::createFromFormat('Y-m-d H:i:s', $fecha_inicio);
  $fecha_inicio = $fecha_inicio->modify('-15 minutes');
  //crear fecha y hora de finalización
  $fecha_final = $fecha_clase." ". $horario['final'][0] . ":". $horario['final'][1] . ":00";
  $fecha_final = DateTime::createFromFormat('Y-m-d H:i:s', $fecha_final);
  //echo $fecha->format('Y-m-d H:i:s');
  $horarios['inicio']['hora'] = $fecha_inicio->format('H');
  $horarios['inicio']['min'] = $fecha_inicio->format('i');
  $horarios['final']['hora'] = $fecha_final->format('H');
  $horarios['final']['min'] = $fecha_final->format('i');
  //print_r($fecha_inicio);
  return $horarios;
}

function transformacionDescripcionCurso($value)
{
  $url_partes = explode("- 528",$value);
  return  trim($url_partes[0]);
}

function imprimirTable($table)
{
  echo "<table
  id='tabla_detalle'
  class='table table-responsive w-100  table-hover'>".
    "<thead>".
      "<tr>".
        "<th>Item</td>".
        "<th>Código Curso</th>".
        "<th>Número</th>".
        "<th>Clase</th>".
        "<th>dia</th>".
        "<th>mes</th>".
        "<th>ano</th>".
        "<th>horai</th>".
        "<th>minutoi</th>".
        "<th>horaf</th>".
        "<th>minutof</th>".
        "<th>descripcion</th>".
        "<th>webex</th>".
      "</tr>".
    "</thead>".
    "<tbody id='tbody_tabla_detalle'>".$table."</tbody>".
  "</table>";

  return true;

}

function agregarRegistroTablaListadosCursos($array)
{

  //print_r($array);

  //Limpiar tabla
  PCO_EjecutarSQLUnaria("TRUNCATE TABLE app_listados_cursos");

  $insertados = 0;
  foreach($array as $linea){
    //echo $linea['codigocurso'] . "-" . $linea['urlimagen'] . "-" . $linea['salaycont']."<br/>";
    //var_dump($linea);
    if( !is_null($linea['urlimagen']) && !is_null($linea['salaycont']) ){
      $sql = "INSERT INTO app_listados_cursos(codigocurso,urlimagen,salaycont,descripcion) ".
        "VALUES('".$linea['codigocurso']."','".$linea['urlimagen']."','".$linea['salaycont']."','".$linea['descripcion']."');";

      PCO_EjecutarSQLUnaria($sql);
      //die();
      $insertados++;
    }

  }

  return "<p>Se han encontrado <b>$insertados</b> registro de cursos.</p>";
}

function agregarRegistroTablaClasesCursosEnVivo($array)
{
  PCO_EjecutarSQLUnaria("TRUNCATE TABLE app_clases_cursos");

  $insertados=0;

  foreach($array as $linea){
    if( !is_null($linea['urlimagen']) && !is_null($linea['webex']) ){
      $sql = "INSERT INTO app_clases_cursos(
        codigocurso,
        numero,
        clase,
        dia,
        mes,
        anio,
        horai,
        minutoi,
        horaf,
        minutof,
        descripcion,
        webex".
        ") ".
        "VALUES(".
        "'".$linea['codigocurso']."',".
        "'".$linea['numero']."',".
        "'".$linea['clase']."',".
        "'".$linea['dia']."',".
        "'".$linea['mes']."',".
        "'".$linea['anio']."',".
        "'".$linea['horai']."',".
        "'".$linea['minutoi']."',".
        "'".$linea['horaf']."',".
        "'".$linea['minutof']."',".
        "'".$linea['descripcion']."',".
        "'".$linea['webex']."'".
        ");";

      PCO_EjecutarSQLUnaria($sql);

      $insertados++;
    }
  }

  return "<p>Se han encontrado <b>$insertados</b> registro de clases.</p>";

}

function encontrarErrores($array, $imprimir = false)
{
  $html_rpta = "<p>";
  foreach($array as $linea){
    //mostrar errores de url imagen del curso
    if(is_null($linea['urlimagen'])){
      $html_rpta .= "Información incompleta en <b>línea " . $linea['linea'] . "</b> Curso:". $linea['nombrecurso'] . " " . $linea['nombrecurso'] . " <br/>[No existe URL de la imagen del curso]<hr/>";
    }
    //mostrar errores de url sala de webex
    if(is_null($linea['webex'])){
      $html_rpta .= "Información incompleta en <b>línea " . $linea['linea'] . "</b> Curso:". $linea['codigocurso'] . " " . $linea['nombrecurso'] . " <br/>[No existe URL de la sala webex]<hr/>";
    }
  }
  $html_rpta .= "</p>";
  if($imprimir) echo $html_rpta;
  return $html_rpta;
}

function NERA_TraerContenidoTabla($nombre_tabla, $campos = "*", $JSON = false)
{
  $sql= "SELECT $campos FROM $nombre_tabla;";

  $results = PCO_EjecutarSQL($sql)->fetchAll(PDO::FETCH_ASSOC);
  //PDO::FETCH_ASSOC PDO::FETCH_CLASS PDO::FETCH_COLUMN
  if($JSON)
    return json_encode($results);
  else
    return $results;
}
