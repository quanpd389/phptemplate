<?php
/**
 * PHPTemplate_Cell
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Cell
 * @copyright
 */
class PHPTemplate_Cell
{
	/**
	 *  セル
	 *
	 *  @var  PHPExcel_Cell
	 */
	private $_cell = NULL;
	/**
	 *  セルが含まれるワークシート
	 *
	 *  @var  PHPExcel_Worksheet
	 */
	private $_worksheet = NULL;
	/**
	 *  セルのアドレス
	 *
	 *  @var  string
	 */
	private $_address = "A1";
	/**
	 *  セルのカラム位置
	 *
	 *  @var  string
	 */
	private $_column = "A";
	/**
	 *  セルの行位置
	 *
	 *  @var  int
	 */
	private $_row = -1;
	/**
	 *  セルのスタイル
	 *
	 *  @var  PHPExcel_Style
	 */
	private $_style = NULL;
	/**
	 *  セルの値
	 *
	 *  @var  string
	 */
	private $_text=NULL;
	/**
	 *  セルの値
	 *
	 *  @var  PHPExcel_RichText
	 */
	private $_richText=NULL;
	/**
	 *  セルの行高さ
	 *
	 *  @var  int
	 */
	private $_rowHeight = -1;
	/**
	 *  セルのカラム幅
	 *
	 *  @var  int
	 */
	private $_columnWidth = -1;
	/**
	 *  ワークシートのデフォルトフォント
	 *
	 *  @var  PHPExcel_Style_Font
	 */
	private $_defaultFont=NULL;
	/**
	 *  セルの結合情報(カラム)
	 *
	 *  @var  array [0]start [1]end
	 */
	private $_mergeInfo=NULL;
	/**
	 *  セルの結合情報(行)
	 *
	 *  @var  bool TRUE:結合開始位置
	 */
	private $_mergeRowStart=FALSE;
	/**
	 *  セルの結合情報(行) 結合行数
	 *
	 *  @var  int
	 */
	private $_mergeRowCnt=0;

	/**
	 *  結合セルの行高さ
	 *
	 *  @var  int
	 */
	private $_mergeRowHeight = -1;
	/**
	 *  結合セルのカラム幅
	 *
	 *  @var  int
	 */
	private $_mergeColumnWidth = -1;

	/**
	 *	Create a new Cell
	 *
	 *	@param	PHPExcel_Cell	$cell
	 */
	public function __construct($cell = NULL, $mergeHeight=-1, $mergeWidth=-1, $mergeRowStart=NULL, $mergeCnt=NULL, $mergeInfo=NULL)
	{
		// Initialise cell value
		$this->_cell = $cell;
		$this->_address = $cell->getCoordinate();
		$this->_column = $cell->getColumn();
		$this->_worksheet = $cell->getWorksheet();
		$this->_style = $cell->getStyle();
		$this->_text = $this->getCellText($this->_cell);
		if (is_object($cell->getValue())) {
			$this->_richText = $cell->getValue();//->getRichTextElements();
		}
		$this->_row = $cell->getRow();

		$this->_rowHeight = $this->_worksheet->getRowDimension($this->_row)->getRowHeight();

//echo "text:".$this->_text."row:".$this->_row.":::".$this->_rowHeight."\n";

		if($this->_rowHeight == -1){
			$this->_rowHeight = PHPTemplate_Util::DEFAULT_ROW_SIZE;
		}

		$this->_columnWidth = $this->_worksheet->getColumnDimension($this->_column)->getWidth();
		if($this->_columnWidth == -1){
			$this->_columnWidth = PHPTemplate_Util::DEFAULT_COLUMN_SIZE;
		}
		$this->_defaultFont = $this->_worksheet->getParent()->getDefaultStyle()->getFont();
		if($mergeRowStart==NULL){
			$this->setMergeInfo();
		}
		else{
			$this->_mergeRowStart = $mergeRowStart;
			$this->_mergeRowCnt = $mergeCnt;
			$this->_mergeInfo = $mergeInfo;
		}
		if($mergeHeight != -1){
			$this->_mergeRowHeight = $mergeHeight;
		}
		if($mergeWidth != -1){
			$this->_mergeColumnWidth = $mergeWidth;
		}
		if($this->_mergeRowHeight == -1){
			$this->_mergeRowHeight = $this->_rowHeight;
		}
		if($this->_mergeColumnWidth == -1){
			$this->_mergeColumnWidth = $this->_columnWidth;
		}


	}

	/**
	 *	Set merge cell info
	 *
	 *
	 */
	private function setMergeInfo(){

		// セル結合の情報を取得
		foreach ($this->_worksheet->getMergeCells() as $mergeCell) {
			$mc = explode(":", $mergeCell);
			$col_s = preg_replace("/[0-9]*/" , "",$mc[0]);
			$col_e = preg_replace("/[0-9]*/" , "",$mc[1]);
			$row_s = ((int)preg_replace("/[A-Z]*/" , "",$mc[0]));
			$row_e = ((int)preg_replace("/[A-Z]*/" , "",$mc[1]));

			if($this->_row == $row_s && $this->_column == $col_s /*&& ($row_e - $row_s) > 0*/){
				$this->_mergeRowStart = TRUE;
				$this->_mergeRowCnt = $row_e - $row_s;
				$merge = Array();
				$merge[]=$col_s;
				$merge[]=$col_e;
				$this->_mergeInfo= $merge;

				// マージセルの高さ幅を取得
				$this->_mergeRowHeight = 0;
				for($i=$row_s; $i<=$row_e; $i++ ){
					$h = $this->_worksheet->getRowDimension($i)->getRowHeight();
					if($h == -1)	$h = PHPTemplate_Util::DEFAULT_ROW_SIZE;
					$this->_mergeRowHeight = $this->_mergeRowHeight + $h;
					//echo $i. ":".$h."mergeRowHeight:".$this->_mergeRowHeight."\n";

				}

				$this->_mergeColumnWidth = 0;
				//echo "index:".PHPExcel_Cell::columnIndexFromString($col_s);
				for($i=(PHPExcel_Cell::columnIndexFromString($col_s)-1); $i<=(PHPExcel_Cell::columnIndexFromString($col_e)-1); $i++ ){
					$c = PHPExcel_Cell::stringFromColumnIndex($i);
					$w = $this->_worksheet->getColumnDimension($c)->getWidth();
					if($w == -1)	$w = PHPTemplate_Util::DEFAULT_COLUMN_SIZE;
					$this->_mergeColumnWidth = $this->_mergeColumnWidth + $w;
					//echo "mergeRowColumnWidth:".$this->_mergeColumnWidth."\n";
				}

				//複数行のマージの場合は結合を解除しておく（次行にマージの情報が残るため）
				//if($this->_mergeRowCnt > 0){
					$this->_worksheet->unmergeCells($mergeCell);
				//}
			}
			/*else{
				// 範囲なら。
				if ($this->_row >= $row_s && $row_e >= $this->_row) {
					if($this->_column >= $col_s && $col_e >= $this->_column){
						$merge = Array();
						$merge[]=$col_s;
						$merge[]=$col_e;
						$this->_mergeInfo= $merge;
					}
				}
			}*/
			//if($this->_row == $row_s && $this->_column == $col_s && ($row_e - $row_s) > 0){
			//	$this->_mergeRowStart = TRUE;
			//	$this->_mergeRowCnt =$row_e - $row_s;
			//}

		}
	}

	/**
	 *	Get text value
	 *
	 *	@return	string
	 */
	public function getCell(){
		return $this->_cell;
	}

	/**
	 *	Get text value
	 *
	 *	@return	string
	 */
	public function getText(){
		return $this->_text;
	}

	/**
	 *	Get bind key
	 *
	 *	@return	string[]
	 */
	public function getBindkey(){
		return $this->getBindkeys($this->getText());
	}

	/**
	 *	Get column address
	 *
	 *	@return	string
	 */
	public function getColumn(){
		return $this->_column;
	}

	/**
	 *	Get worksheet
	 *
	 *	@return	PHPExcel_Worksheet
	 */
	public function getWorkSheet(){
		return $this->_worksheet;
	}

	/**
	 *	Get cell style
	 *
	 *	@return	PHPExcel_Style
	 */
	public function getStyle(){
		return $this->_style;
	}

	/**
	 *	Get row number
	 *
	 *	@return	int
	 */
	public function getRow(){
		return $this->_row;
	}
	/**
	 *	Get row row height
	 *
	 *	@return	int
	 */
	public function getRowHeight(){
		return $this->_rowHeight;
	}
	/**
	 *	Get row column width
	 *
	 *	@return	int
	 */
	public function getColumnWidth(){
		return $this->_columnWidth;
	}
	/**
	 *	Get merge column info
	 *
	 *	@return	string[]
	 */
	public function getMergeInfo() {
		return $this->_mergeInfo;
	}
	/**
	 *	Get merge Row
	 *
	 *	@return	bool
	 */
	public function getMergeRowStart() {
		return $this->_mergeRowStart;
	}
	/**
	 *	Get merge row count
	 *
	 *	@return	int
	 */
	public function getMergeRowCnt() {
		return $this->_mergeRowCnt;
	}
	/**
	 *	Set row number
	 *
	 *	@param	int  $r
	 */
	public function setRow($r){
		$this->_row = $r;
	}
	/**
	 *	Get merge row row height
	 *
	 *	@return	int
	 */
	public function getMergeRowHeight(){
		return $this->_mergeRowHeight;
	}
	/**
	 *	Get merge column width
	 *
	 *	@return	int
	 */
	public function getMergeColumnWidth(){
		return $this->_mergeColumnWidth;
	}

	/**
	 *	バインド変数を設定
	 *
	 *	@param	array	  			$data
	 *	@param	int  				$r
	 *	@param	PHPTemplate_Page  	$page
	 */
	public function bindData($data, $r=NULL, $page)
	{
		//スタイルコピー
		$style = $this->getStyle();
		$this->_worksheet->duplicateStyle(clone $style, $this->_column.$r);

		#linkか？
		$text = $this->getText();
		$type = PHPTemplate_Util::getControlKind($text, $this->getColumn());
		if(		$type==PHPTemplate_Util::ROW_TYPE_LINK_URL
			||	$type==PHPTemplate_Util::ROW_TYPE_LINK_FILE
			||	$type==PHPTemplate_Util::ROW_TYPE_LINK_EMAIL
			||	$type==PHPTemplate_Util::ROW_TYPE_LINK_THIS){

			$pos = strpos($text, " ");
			//#xxxを削除
			if($pos!== false){
				//OK
				$text = trim(substr($text, $pos));
			}
			$url=PHPTemplate_Util::getParamater($text, "link");
			$dsp=PHPTemplate_Util::getParamater($text, "text");
			//バインドもする
			$keys = self::getBindkeys($url);
			if($keys != NULL){
				foreach ($keys as $value) {
					$key = trim($value,'\$\{\}');
					//バインドデータ取得
					$d = PHPTemplate_Util::getBindData($key, $data, $r, $page->getPagenum());
					$url=str_replace($value, $d, $url);
				}
			}
			$keys = self::getBindkeys($dsp);
			if($keys != NULL){
				foreach ($keys as $value) {
					$key = trim($value,'\$\{\}');
					//バインドデータ取得
					$d = PHPTemplate_Util::getBindData($key, $data, $r, $page->getPagenum());
					$dsp=str_replace($value, $d, $dsp);
				}
			}
			if($type==PHPTemplate_Util::ROW_TYPE_LINK_FILE){
				$url="file:///".$url;
			}
			elseif ($type==PHPTemplate_Util::ROW_TYPE_LINK_EMAIL){
				$url="mailto:".$url;
			}
			elseif ($type==PHPTemplate_Util::ROW_TYPE_LINK_THIS){
				$url="sheet://".$url;
			}
			$link = new PHPExcel_Cell_Hyperlink($url, $dsp);
			$this->_cell = $this->_worksheet->getCell($this->_column.$r);
			$this->_cell->setHyperlink($link);
			$this->setCellValue($dsp, $r);
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_IMG){//画像か？
			$path=PHPTemplate_Util::getParamater($text, "file");
			//バインドもする
			$keys = self::getBindkeys($path);
			if($keys != NULL){
				foreach ($keys as $value) {
					$key = trim($value,'\$\{\}');
					//バインドデータ取得
					$d = PHPTemplate_Util::getBindData($key, $data, $r, $page->getPagenum());
					$path=str_replace($value, $d, $path);
				}
			}
			$this->setCellValue("", $r);

			///画像用のオプジェクト作成
			if(file_exists($path)){
				$objDrawing = new PHPExcel_Worksheet_Drawing();
				$objDrawing->setPath($path);///貼り付ける画像のパスを指定
				//フォント情報と列幅からピクセルを取得
				//$colPixelWidth = PHPExcel_Shared_Drawing::cellDimensionToPixels($this->_columnWidth, $this->_defaultFont);
				//$rowPixelHeight = PHPExcel_Shared_Drawing::pointsToPixels($this->_rowHeight);
				$colPixelWidth = PHPExcel_Shared_Drawing::cellDimensionToPixels($this->_mergeColumnWidth, $this->_defaultFont);
				$rowPixelHeight = PHPExcel_Shared_Drawing::pointsToPixels($this->_mergeRowHeight);
				//echo "draw:".$this->_mergeColumnWidth.":".$this->_mergeRowHeight."\n";
				//$objDrawing->setWidth($colPixelWidth);
				//$objDrawing->setHeight($rowPixelHeight);
				$objDrawing->setWidthAndHeight($colPixelWidth-2, $rowPixelHeight-2);
				//センタリング
				$objDrawing->setOffsetX(($colPixelWidth-$objDrawing->getWidth())/2);
				$objDrawing->setOffsetY(($rowPixelHeight-$objDrawing->getHeight())/2);

				$objDrawing->setCoordinates($this->_column.(string)$r);///位置
				$objDrawing->setWorksheet($this->_worksheet);
			}
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_SUSPEND){
			// バインド変数のまま #suspendを外す
			$val = trim(substr($text, strlen("#suspend")));
			$this->setCellValue($val, $r);
			$this->_text=$val;
		}
		else{
			$keys = $this->getBindkey();
			if($keys == NULL){
				if($this->_richText != NULL){
					//echo ("get_class:".gettype($this->_richText)."\n");
					$this->setCellValue($this->_richText, $r);
				}
				else{
					$this->setCellValue($text, $r);
				}
				return;
			}
			foreach ($keys as $value) {
				$key = trim($value,'\$\{\}');

				//バインドデータ取得
				$d = PHPTemplate_Util::getBindData($key, $data, $r, $page->getPagenum());
				if(gettype($d)!='string' && count($keys)==1 && $value===$text){
					$this->setCellValue($d, $r, $page->getFont());
					return;
				}
				else{
					$text = str_replace($value, $d, $text);
					if($this->_richText!=NULL){
						$rtfCell = $this->_richText->getRichTextElements();
						foreach ($rtfCell as $v) {
							$v->setText(str_replace($value, $d, $v->getText()));
						}
						$this->_richText->setRichTextElements($rtfCell);
					}

				}
			}
			if($this->_richText!=NULL){
				$this->setCellValue($this->_richText, $r);
			}
			else{
				$this->setCellValue($text, $r, $page->getFont());
			}
		}
	}

	/**
	 *	セルにデータをセットする
	 *
	 *	@param	array	  			$data
	 *	@param	int  				$r
	 *	@param	PHPTemplate_Font[] 	$font
	 */
	private function setCellValue($data, $r, $font=NULL){
		//DateTime型のデータ
		if(gettype($data)=='object' && get_class($data)=='DateTime'){
			$this->_cell = $this->_worksheet->getCell($this->_column.$r);
			$datetime1 = new DateTime('1900-01-01');//excel基準日
			$datetime2 = $data;
			$interval = $datetime2->diff($datetime1);
			$serial = (float)($interval->format('%a'))+2;
			$totalTime = (float)($interval->format('%h'))*60*60+(float)($interval->format('%i'))*60+(float)($interval->format('%s'));
			$totalTime = (float)$totalTime/(float)(24*60*60);
			$serial = $serial+$totalTime;
			$this->_cell->setValueExplicit($serial, PHPExcel_Cell_DataType::TYPE_NUMERIC);
			return;

		}
		elseif(gettype($data)=='object' && get_class($data)=='PHPExcel_RichText')
		{
			$this->_cell = $this->_worksheet->getCell($this->_column.$r);
			$this->_cell->setValue($data);
			return;
		}
		// 文字列のデータ
		$textArray = $this->getTagText($data);
		if(count($textArray) <= 1){
			$this->_cell = $this->_worksheet->getCell($this->_column.$r);
			$this->_cell->setValue($data);
			return;
		}

		$objRichText = new PHPExcel_RichText();
		foreach ($textArray as $text){
			if($text->getTagName()==NULL){
				//$objRichText->createText($text->_text);
				$objFont = $objRichText->createTextRun($text->getText());
				$objFont->setFont(clone $this->getStyle()->getFont());
			}
			else{
				foreach ($font as $f){

					if(($f->getTagname())===($text->getTagName())){
						$objFont = $objRichText->createTextRun($text->getText());
						$objFont->setFont(clone $this->getStyle()->getFont());
						if($f->getColor() != -1){
							$color = new PHPExcel_Style_Color();
							$color->setRGB($f->getColor());
							$objFont->getFont()->setColor($color);
						}
						if($f->getBold() != NULL){
							$objFont->getFont()->setBold($f->getBold());
						}
						if($f->getItalic() != NULL){
							$objFont->getFont()->setItalic($f->getItalic());
						}
						if($f->getUnderline() != NULL){
							$objFont->getFont()->setUnderline($f->getUnderline());
						}
						break;
					}
				}
			}
		}
		$this->_cell = $this->_worksheet->getCell($this->_column.$r);
		$this->_cell->setValue($objRichText);

	}

	/**
	 *	タグ設定されている文字列を取得する
	 *
	 *	@param	string	  			$str
	 *	@return array
	 */
	public function getTagText($str){
		$data = $str;
		$textArray = array();
		$posS = strpos($data, "<");
		$posE = strpos($data, ">");
		if($posS===FALSE || $posE===FALSE){
			$info = new PHPTemplate_TextInfo(NULL, $data);
			$textArray[]=$info;
			return $textArray;
		}
		while($posS!==FALSE && $posE!==FALSE){
			if($posS!=0){
				$info = new PHPTemplate_TextInfo(NULL, substr($data, 0, $posS));
				$textArray[]=$info;
			}
			$info = new PHPTemplate_TextInfo(substr($data, $posS+1, $posE-($posS+1)), NULL);

			$data = substr($data, $posE+1);
			$posS = strpos($data, "</");
			$posE = strpos($data, ">");
			if($posS===FALSE || $posE===FALSE){
				$info->setText($data);
				$textArray[]=$info;
				break;
			}
			$info->setText(substr($data, 0, $posS));
			$textArray[]=$info;
			$data = substr($data, $posE+1);
			$posS = strpos($data, "<");
			$posE = strpos($data, ">");
		}
		if(strlen($data) > 0){
			$info = new PHPTemplate_TextInfo(NULL, $data);
			$textArray[]=$info;
		}
		return $textArray;



	}

	/**
	 *	バインド変数か判定　バインド変数ならキーを返す
	 *
	 *	@param	string	  			$text
	 *	@return array
	 */
	public static function getBindkeys($text) {
		if(is_null($text))
			return NULL;
		if(strlen($text) == 0)
			return NULL;
		// ${}で囲まれている文字列を含むかチェック
		$str = trim($text);
		if(preg_match_all('/\$\{.+?\}/', $str, $out) > 0){
			//バインド変数あり
			return $out[0];
		}
		else{
			return NULL;
		}
	}

	/**
	 *	セル内の文字列を返す
	 *
	 *	@param	PHPExcel_Cell		$objCell
	 *	@return string
	 */
	public static function getCellText($objCell = null){
		if (is_null($objCell)) {
			return NULL;
		}
		$txtCell = "";
		$valueCell = $objCell->getValue();
		if (is_object($valueCell)) {
			$rtfCell = $valueCell->getRichTextElements();
			$txtParts = array();
			foreach ($rtfCell as $v) {
				$txtParts[] = $v->getText();
			}
			$txtCell = implode("", $txtParts);
		} else {
			if (!empty($valueCell)) {
				$txtCell = $valueCell;
			}
		}
		return $txtCell;
	}


}
