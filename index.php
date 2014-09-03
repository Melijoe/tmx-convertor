<?php
	defined('APPLICATION_PATH')         || define('APPLICATION_PATH',       realpath(__DIR__));
	defined('DOCROOT_PATH')             || define('DOCROOT_PATH',           realpath(APPLICATION_PATH."/.."));

	include_once(__DIR__."/PHPExcel/PHPExcel.php");

	class Cell
	{
	    const DEFAULT_BG_COLOR = "FFFFFF";

	    /* Horizontal alignment styles */
	    const HORIZONTAL_LEFT   = 'left';
	    const HORIZONTAL_RIGHT  = 'right';
	    const HORIZONTAL_CENTER = 'center';

	    /* Vertical alignment styles */
	    const VERTICAL_BOTTOM   = 'bottom';
	    const VERTICAL_TOP      = 'top';
	    const VERTICAL_CENTER   = 'center';
	    const VERTICAL_JUSTIFY  = 'justify';

	    /**
	     * @var bool
	     */
	    protected $bLocked = true;

	    /**
	     * @var string
	     */
	    protected $sBackgroundColor = self::DEFAULT_BG_COLOR;

	    /**
	     * @var mixed
	     */
	    protected $mValue;

	    /**
	     * @var string
	     */
	    protected $sHorizontalAlign = self::HORIZONTAL_LEFT;

	    /**
	     * @var string
	     */
	    protected $sVerticalAlign = self::VERTICAL_BOTTOM;

	    /**
	     * Lock a cell
	     */
	    public function lock() {
	        $this->bLocked = true;
	    }

	    /**
	     * Unlock a cell
	     */
	    public function unlock() {
	        $this->bLocked = false;
	    }

	    /**
	     * @return boolean
	     */
	    public function isLocked() {
	        return $this->bLocked;
	    }

	    /**
	     * @param string $sBackgroundColor
	     */
	    public function setBackgroundColor($sBackgroundColor) {
	        $this->sBackgroundColor = $sBackgroundColor;
	    }

	    /**
	     * @return string
	     */
	    public function getBackgroundColor() {
	        return $this->sBackgroundColor;
	    }

	    /**
	     * @param boolean $bLocked
	     */
	    public function setBLocked($bLocked) {
	        $this->bLocked = $bLocked;
	    }

	    /**
	     * @return boolean
	     */
	    public function getBLocked() {
	        return $this->bLocked;
	    }

	    /**
	     * @param mixed $mValue
	     */
	    public function setValue($mValue) {
	        $this->mValue = $mValue;
	    }

	    /**
	     * @return mixed
	     */
	    public function getValue() {
	        return $this->mValue;
	    }

	    /**
	     * @param string $sHorizontalAlign
	     */
	    public function setHorizontalAlign($sHorizontalAlign) {
	        $this->sHorizontalAlign = $sHorizontalAlign;
	    }

	    /**
	     * @return string
	     */
	    public function getHorizontalAlign() {
	        return $this->sHorizontalAlign;
	    }

	    /**
	     * @param string $sVerticalAlign
	     */
	    public function setVerticalAlign($sVerticalAlign) {
	        $this->sVerticalAlign = $sVerticalAlign;
	    }

	    /**
	     * @return string
	     */
	    public function getVerticalAlign() {
	        return $this->sVerticalAlign;
	    }
	}

	class HeaderCell extends Cell
	{
	    /**
	     * @var bool
	     */
	    protected $bLocked = true;

	    /**
	     * @var string
	     */
	    protected $sBackgroundColor = "F28A8C";

	    /**
	     * @var string
	     */
	    protected $sHorizontalAlign = self::HORIZONTAL_CENTER;

	    /**
	     * @var string
	     */
	    protected $sVerticalAlign = self::VERTICAL_CENTER;

	    /**
	     * @var bool
	     */
	    protected $bHidden = false;

	    /**
	     * Hide a column
	     */
	    public function hide() {
	        $this->bHidden = true;
	    }

	    /**
	     * Show a column
	     */
	    public function show() {
	        $this->bHidden = false;
	    }

	    /**
	     * @return boolean
	     */
	    public function isHidden() {
	        return $this->bHidden;
	    }
	}

	class XlsWriter
	{
	    /**
	     * @var \PHPExcel
	     */
	    protected $_objPHPExcel;
	    /**
	     * @var \PHPExcel_Worksheet
	     */
	    protected $_oWorksheet;
	    /**
	     * @var string
	     */
	    protected $_sFilename;
	    /**
	     * @var int
	     */
	    protected $_currentRow = 1;
	    /**
	     * @var int
	     */
	    protected $_currentColumn = 0;

	    /**
	     * @param string $sFilename
	     * @param bool   $bProtected
	     */
	    public function __construct($sFilename, $bProtected = true) {
	        $this->_sFilename = $sFilename;
	        $this->_objPHPExcel = new \PHPExcel();

	        $this->_oWorksheet = $this->_objPHPExcel->getActiveSheet();
	        $this->_oWorksheet->getProtection()->setSheet($bProtected);
	    }

	    /**
	     * @param string $sKey
	     * @param string $sValue
	     */
	    public function setMetadata($sKey, $sValue) {
	        $this->_objPHPExcel->getProperties()->setCustomProperty($sKey, $sValue);
	    }

	    private function setCell(Cell $oCell, \PHPExcel_Cell $oFileCell) {
	        $oFileCell->setValue($oCell->getValue());
	        $oFileCellStyle = $this->_oWorksheet->getStyle($oFileCell->getCoordinate());

	        /**
	         * Lock of the cell
	         */
	        if ($oCell->isLocked() === false) {
	            $this->_oWorksheet->getStyle($oFileCell->getCoordinate())->getProtection()->setLocked(\PHPExcel_Style_Protection::PROTECTION_UNPROTECTED);
	        } else {
	            $oFileCellStyle->applyFromArray(array(
	                'fill' => array(
	                    'type' => \PHPExcel_Style_Fill::FILL_SOLID,
	                    'color' => array('rgb' => 'DEDEDE')
	                )
	            ));
	        }

	        /**
	         * Style
	         */
	        $aStyleArray = array(
	            'borders' => array(
	                'outline' => array(
	                    'style' => \PHPExcel_Style_Border::BORDER_HAIR,
	                    'color' => array('rgb' => '000000'),
	                ),
	            ),
	        );
	        if ($oCell->getBackgroundColor() != Cell::DEFAULT_BG_COLOR) {
	            $aStyleArray['fill'] = array(
	                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
	                'color' => array('rgb' => $oCell->getBackgroundColor())
	            );
	        }
	        $oFileCellStyle->applyFromArray($aStyleArray);
	        if ($oCell->getHorizontalAlign() != Cell::HORIZONTAL_LEFT) {
	            $oFileCellStyle->getAlignment()->setHorizontal($oCell->getHorizontalAlign());
	        }
	        if ($oCell->getVerticalAlign() != Cell::VERTICAL_BOTTOM) {
	            $oFileCellStyle->getAlignment()->setVertical($oCell->getVerticalAlign());
	        }

	    }

	    /**
	     * @param HeaderCell[] $aCells
	     */
	    public function setHeader($aCells) {

	        foreach ($aCells as $oCell) {
	            $oFileCell = $this->_oWorksheet->getCellByColumnAndRow($this->_currentColumn, $this->_currentRow);

	            $this->setCell($oCell, $oFileCell);

	            if ($oCell->isHidden()) {
	                $this->_oWorksheet->getColumnDimensionByColumn($this->_currentColumn)->setVisible(false);
	            } else {
	                $this->_oWorksheet->getColumnDimensionByColumn($this->_currentColumn)->setAutoSize(true);
	            }

	            $this->_currentColumn++;
	        }
	        $this->_currentRow++;
	        $this->_currentColumn = 0;
	    }

	    /**
	     * @param Cell[] $aCells
	     */
	    public function write($aCells) {

	        foreach ($aCells as $oCell) {
	            $oFileCell = $this->_oWorksheet->getCellByColumnAndRow($this->_currentColumn, $this->_currentRow);

	            $this->setCell($oCell, $oFileCell);

	            $this->_currentColumn++;
	        }
	        $this->_currentRow++;
	        $this->_currentColumn = 0;
	    }

	    /**
	     * @return bool
	     */
	    public function save() {
	        try {
	            $objWriter = \PHPExcel_IOFactory::createWriter($this->_objPHPExcel, 'Excel2007');
	            $objWriter->save($this->_sFilename);
	        } catch (\Exception $e) {
	            return false;
	        }
	        return true;
	    }
	}

	if(isset($_FILES["file"])) {
		$target_Path = "tmp/";
		$target_Path = $target_Path.basename( $_FILES['userFile']['name'] );

		move_uploaded_file( $_FILES['file']['tmp_name'], $target_Path . $_FILES['file']['name'] );

	    $sFilePath = $target_Path . $_FILES['file']['name'];

	    // Lecture du fichier tmx
	    $oXml = simplexml_load_file($sFilePath);
	    if (isset($oXml->body->tu)) {
	        // Construction d'un array contenant toutes les valeur "Source" et "Translation" du fichier .tmx
	        $aCsv = array();
	        foreach ($oXml->body->tu as $oLine) {
	            $sColonne1 = isset($oLine->tuv[0]->seg) ? $oLine->tuv[0]->seg : "";
	            $sColonne2 = isset($oLine->tuv[1]->seg) ? $oLine->tuv[1]->seg : "";
	            $aCsv[] = array($sColonne1, $sColonne2);
	        }

	        // Création du fichier Excel
	        $oWriter = new XlsWriter($target_Path . str_replace(".tmx", ".xlsx", $_FILES['file']['name']));
	        $oWriter->setMetadata("type", "contexts");

	        // Création du header du fichier
	        $aHeader = array();

	        // Colonne 1
	        $oHeaderCell = new HeaderCell();
	        $oHeaderCell->setValue("Source");
	        $aHeader[] = $oHeaderCell;

	        // Colonne 2
	        $oHeaderCell = new HeaderCell();
	        $oHeaderCell->setValue("Translation");
	        $aHeader[] = $oHeaderCell;

	        $oWriter->setHeader($aHeader);

	        foreach ($aCsv as $fields) {
	            $aCells = array();

	            $oCell = new Cell();
	            $oCell->setValue($fields[0]);
	            $aCells[] = $oCell;

	            $oCell = new Cell();
	            $oCell->setValue($fields[1]);
	            $aCells[] = $oCell;

	            $oWriter->write($aCells);
	        }
	        $oWriter->save();
	    }
		
		$filename = str_replace(".tmx", ".xlsx", $_FILES['file']['name']);
		$file = $target_Path . $filename;

		header('Content-disposition: attachment; filename="' . $filename . '"');
		header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Length: ' . filesize($file));
		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		
		ob_clean();
		flush();

		readfile($file);
		//On supprime tout les fichiers dans uploaded
		$dir = 'your/directory/';
		foreach(glob($target_Path.'*.*') as $v){
		    unlink($v);
		}

		exit;
	} else { 
?>

	<html lang="en">
		<head>
		    <meta charset="utf-8">
		    <meta http-equiv="X-UA-Compatible" content="IE=edge">
		    <meta name="viewport" content="width=device-width, initial-scale=1">
		    <meta name="description" content="">
		    <meta name="author" content="">
		    <title>Parser TMX into Excel.</title>
		    <!-- Bootstrap core CSS -->
		    <link href="css/bootstrap.min.css" rel="stylesheet">
		    <script type="text/javascript" src="js/jquery-2.1.1.min.js"> </script>
		    <script type="text/javascript" src="js/bootstrap.min.js"> </script>
		    <style type="text/css">
				html, body {
					height: 100%;
					background-color: #333;
				}
				body {
					color: #fff;
					text-align: center;
					text-shadow: 0 1px 3px rgba(0,0,0,.5);ackground-color: #eee;
				}
				.site-wrapper {
					display: table;
					width: 100%;
					height: 100%;
					min-height: 100%;
					-webkit-box-shadow: inset 0 0 100px rgba(0,0,0,.5);
					box-shadow: inset 0 0 100px rgba(0,0,0,.5);
				}
				.site-wrapper-inner {
					display: table-cell;
					vertical-align: middle;
				}
				.masthead, .mastfoot, .cover-container {
					width: 700px;
				}
				.cover-container {
					margin-right: auto;
					margin-left: auto;
				}
				.input-xlarge {
					color: #000;
				}
		    </style>
		</head>
		<body>
			<div class="site-wrapper">
	      		<div class="site-wrapper-inner">
	        		<div class="cover-container">
	          			<div class="inner cover">
	            			<h1 class="cover-heading">Parser TMX into Excel.</h1>
							<p class="lead">
								<label>Choose A File</label>
								<div class="input-append">
									<form enctype="multipart/form-data" method="post">
										<!-- This input is here purely for cosmetic reasons. The actual file is uploaded from the hidden input box !-->
										<input type="file" name="file" style="visibility:hidden;" id="input_file">
										<input type="text" id="subfile" class="input-xlarge">
										<input type="button" class="btn btn-primary" value="Browse" onclick="$('#input_file').click();" />
										&nbsp;|&nbsp;
										<input type="submit" class="btn btn-success" value="Convert" name="convert">
									</form>
								</div> 
							</p>
	          			</div>
	        		</div>
	      		</div>
	    	</div>
		</body>
	</html>

	<script>
		$(document).ready(function(){
			// This is the simple bit of jquery to duplicate the hidden field to subfile
			$(document).on("change", "#input_file", function() {
				$('#subfile').val($(this).val());
			}); 
		});
	</script>

<?php } ?>