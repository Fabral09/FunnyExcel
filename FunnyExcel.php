<?php

class FunnyExcel
{
    /*
     *  Costruttore di classe  
     */
    public function __construct()
	{
        $this->lib_dir = "./classes/";
        require_once $this->lib_dir . "PHPExcel.php";
        $this->inputFileType = 'Excel2007';
        $this->objExcel = new PHPExcel();
        $this->usedColumns = array();       // Array delle colonne usate nel Excel
    }

    /*
     *  Imposta il nome del file da generare
     */
    public function SetFileName( $fileName, $indexSheet = 0 )
	{
        $this->fileName = $fileName;
        $this->activeSheet = $indexSheet;
        $this->SetActiveSheet( $this->activeSheet );     // Imposto il foglio attivo
    }

    /*
     * Imposta le informazioni aggiuntive per
     * il file Excel 
     */
    public function SetAdditionalInfo( $creatorName, $lastModifiedBy, $documentTitle, $documentSubject, $documentDescription )
    {
        // Imposto le proprietà del file
        $this->objExcel->getProperties()->setCreator( $creatorName );
        $this->objExcel->getProperties()->setLastModifiedBy( $lastModifiedBy );
        $this->objExcel->getProperties()->setTitle( $documentTitle );
        $this->objExcel->getProperties()->setSubject( $documentSubject );
        $this->objExcel->getProperties()->setDescription( $documentDescription );
    }    
    
    /*
     *  Imposta il foglio Excel da utilizzare nella classe
     */
    public function SetActiveSheet( $indexSheet )
	{
        $this->activeSheet = $indexSheet;
        $this->objExcel->setActiveSheetIndex( $this->activeSheet );     // Imposto il foglio attivo
    }

    /*
     *  Imposta l'header ( riga di intestazione ) del file excel
     */
    public function SetConfig( $sheetName, $configArray, $cellBackgroundColor, $cellTextColor )
	{
        $this->objExcel->getActiveSheet()->setTitle( $sheetName );      // Setto il titolo del foglio excel
        // Scorro l'array delle colonne da utilizzare
        foreach ( $configArray as $key => $value )
		{
            // Applico la formattazione alla prima riga in ogni colonna da utilizzare
            $this->objExcel->getActiveSheet()->getStyle( $key . "1" )->applyFromArray( 
                                                                        array( 
                                                                                'font' => array( 'color' => array( 'rgb' => $cellTextColor ), 'bold' => true, ),
                                                                                'alignment' => array( 'wrap' => true, 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, 'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER ),
                                                                                'fill' => array( 'type' => PHPExcel_Style_Fill::FILL_SOLID, 'color' => array( 'rgb' => $cellBackgroundColor ) ) ) );
            // Aggiungo l'intestazione nella prima riga con le colonne specificate
            $this->objExcel->getActiveSheet()->SetCellValue( $key . "1", $value );
            // Memorizzo la colonna attuale nell'elenco delle colonne da usare nella memorizzazione dati
            $this->usedColumns[] = $key;
        }
    }
    
    /*
     * Setta l'allineamento [ 0 = a sinistra, 1 = al centro, 2 = a destra ], su una
     * colonna per un determinato numero di righe
     * 
     * DA TERMINARE L'IMPLEMENTAZIONE
     */
    
    /*
    public function SetAlignmentToColumn( $column, $fromInitialRow, $toFinalRow, $alignement )
    {
        for ( $row = $fromInitialRow; $row <= $toFinalRow; $row++ )
        {
            switch( $alignement )
            {
                case 0:
                    $this->objExcel->getActiveSheet()->getStyle( $column . $row )->applyFromArray( array( 'alignment' => array( 'wrap' => true, 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT ) ) );
                    break;
                case 1:
                    $this->objExcel->getActiveSheet()->getStyle( $column . $row )->applyFromArray( array( 'alignment' => array( 'wrap' => true, 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER ) ) );
                    break;
                case 2:
                    $this->objExcel->getActiveSheet()->getStyle( $column . $row )->applyFromArray( array( 'alignment' => array( 'wrap' => true, 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT ) ) );
                    break;
                default:
                    break;
            }
       }  
    }*/
    
    /*
     * Inserisce i dati dal recordset nell'oggetto
     * creando virtualmente il foglio Excel
     */
    public function SetData( $result, $actualRow )
	{
        foreach ( $result as $row )
		{
            $columnIndex = 0;           // Imposto la colonna iniziale
            foreach ( $row as $item )
			{
                $this->objExcel->getActiveSheet()->SetCellValue( $this->usedColumns[$columnIndex] . $actualRow, $item );
                $columnIndex++;     // Mi sposto sulla colonna successiva
            }
            $actualRow++;   // Mi sposto sulla riga successiva
        }
    }

    /*
     * Imposta le protezioni sul file Excel 
     */
    public function SetSecurity( $filePassword, $sheetPassword )
    {
        $this->objExcel->getSecurity()->setLockWindows( true );
        $this->objExcel->getSecurity()->setLockStructure( true );
        $this->objExcel->getSecurity()->setWorkbookPassword( $filePassword );
        $this->objExcel->getActiveSheet()->getProtection()->setPassword( $sheetPassword );
        $this->objExcel->getActiveSheet()->getProtection()->setSheet( true );
        $this->objExcel->getActiveSheet()->getProtection()->setSort( true );
        $this->objExcel->getActiveSheet()->getProtection()->setInsertRows( true );
        $this->objExcel->getActiveSheet()->getProtection()->setFormatCells( true );
    }
    
    /*
     *  Scrive il file Excel
     */
    public function WriteFile()
    {
        // Costruisco il nome del file in modo univoco
        $date = date( 'd-m-Y-His' );
        $completeNameWithDate = $this->fileName . $date . ".xlsx";
        // Genero il writer
        $objWriter = PHPExcel_IOFactory::createWriter( $this->objExcel, $this->inputFileType );
        // Invio gli header al browser, in modo da farmi partire il download del file
        header( 'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
        header( 'Content-Disposition: attachment;filename="' . $completeNameWithDate . '"' );
        header( 'Cache-Control: max-age=0' );
        $objWriter->save( 'php://output' );   // Invio lo stream dati
        $this->objExcel->disconnectWorksheets();    // Chiudo i fogli Excel dell'oggetto
        die();
    }
    
    /*
     *  Scrive il file Excel su disco
     */
    public function WriteFileOnDisk( $directory = "." )
    {
        // Costruisco il nome del file in modo univoco
        $date = date( 'd-m-Y-His' );
        $completeNameWithDate = $this->fileName . $date . ".xlsx";
        // Genero il writer
        $objWriter = PHPExcel_IOFactory::createWriter( $this->objExcel, $this->inputFileType );
        $objWriter->save( $directory . "/" . $completeNameWithDate );   // Salvo il file
        $this->objExcel->disconnectWorksheets();    // Chiudo i fogli Excel dell'oggetto
        die();
    }
	
    /*
     * Membri privati
     */
    private $objExcel;
    private $fileName;
    private $activeSheet;
    private $usedColumns;
}

?>
