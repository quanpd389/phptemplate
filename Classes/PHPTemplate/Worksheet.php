<?php
/**
 * PHPTemplate_Worksheet
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Worksheet
 * @copyright
 */
class PHPTemplate_Worksheet
{
	/**
	 * 行リスト
	 * @var	array
	 */
	private $_rowArray = array();

	/**
	 * ワークシート
	 * @var	PHPExcel_Sheet
	 */
	private $_sheet = NULL;

	/**
	 * パラメータリスト
	 * @var	array
	 */
	private $_paramArray = array();

	/**
	 * ページ管理
	 * @var	PHPTemplate_Page
	 */
	private $_page=NULL;


	/**
	 *
	 */
	//private $_rowDimensions=array();

	/**
	 *	Create a new Sheet
	 *
	 *	@param	PHPExcel_Sheet	$sheet
	 */
	public function __construct($sheet = NULL)
	{
		$this->_sheet = $sheet;
		$this->_page = new PHPTemplate_Page();
	}

	/**
	 *rowArrayへのデータ追加
	 *
	 *@param	mixed		$obj
	 * */
	public function addRow($obj)
	{
		$this->_rowArray[] = $obj;
	}

	/**
	 *paramArrayへのデータ追加
	 *
	 *@param	string		$key
	 *@param	string		$val
	 * */
	public function addParam($key, $val)
	{
		$this->_paramArray[$key] = $val;
	}

	/**
	 *paramArrayからのデータ取得
	 *
	 *@param	string		$key
	 *@return	string
	 * */
	public function getParam($key)
	{
		return $this->_paramArray[$key];
	}

	/**
	 * シートを分析してデータをセット
	 *
	 * @param	array	$data		バインドデータ
	*/
	public function analyzeSheet($data){

		$this->_paramArray = $data;

		//print_r($this->_sheet->getRowDimensions());
		//$this->_rowDimensions = $this->_sheet->getRowDimensions();

		$rowMax = $this->_sheet->getHighestRow();
		$colMax = PHPExcel_Cell::columnIndexFromString($this->_sheet->getHighestColumn());
		for($r=1; $r<=$rowMax; $r++) {
			$row = new PHPTemplate_Row();
			$rowAdd = TRUE;
			for($c=0; $c<=$colMax; $c++) {
				$cell = $this->_sheet->getCellByColumnAndRow($c, $r);
				$objCell = new PHPTemplate_Cell($cell, -1, -1);
				$type = $this->getBlock($objCell, $r, $row, $colMax, $rowMax, $this);

				if($type == PHPTemplate_Util::ROW_TYPE_NONE){
				}
				else{
					$rowAdd = FALSE;
					break;
				}

			}
			if($rowAdd){
				$this->_rowArray[] = $row;
			}
		}

		for($r=1; $r<=$rowMax; $r++) {
			$this->_sheet = $this->_sheet->removeRow(1);
		}
		//$rowMax = $this->_sheet->getHighestRow();

		//データバインド
		$count = 1;
		foreach($this->_rowArray as $rw){
			$rw->expand($this->_paramArray, $count, $this->_page);
		}
		//シート名バインド
		$sheetName = $this->_sheet->getTitle();
		$keys = PHPTemplate_Cell::getBindkeys($sheetName);
		if($keys == NULL){
			return;
		}
		foreach ($keys as $value) {
			$key = trim($value,'\$\{\}');
			//バインドデータ取得
			$d = PHPTemplate_Util::getBindData($key, $data, NULL, NULL);
			$sheetName = str_replace($value, $d, $sheetName);
		}
		$this->_sheet->setTitle($sheetName);
	}


	/**
	 * 制御ブロックを取得する
	 *
	 * @param	PHPTemplate_Cell		$objCell
	 * @param	int						$r
	 * @param	PHPTemplate_Row			$row
	 * @param	int						$colMax
	 * @param	int						$rowMax
	 * @param	array					$objRows
	*/
	public function getBlock($objCell, &$r, $row, $colMax, $rowMax, $objRows){

		$type = PHPTemplate_Util::getControlKind($objCell->getText(), $objCell->getColumn());

		if($type == PHPTemplate_Util::ROW_TYPE_NONE
		|| $type == PHPTemplate_Util::ROW_TYPE_LINK_EMAIL
		|| $type == PHPTemplate_Util::ROW_TYPE_LINK_FILE
		|| $type == PHPTemplate_Util::ROW_TYPE_LINK_THIS
		|| $type == PHPTemplate_Util::ROW_TYPE_LINK_URL
		|| $type == PHPTemplate_Util::ROW_TYPE_IMG
		|| $type == PHPTemplate_Util::ROW_TYPE_SUSPEND){
			$row->addCell($objCell);
			return PHPTemplate_Util::ROW_TYPE_NONE;
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_PAGEBREAK){
			$obj = new PHPTemplate_PageBreak($objCell->getText(), $this->_sheet);
			$objRows->addRow($obj);
			//$this->_sheet->removeRow($r);
			//$r--;
			return $type;
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_END){
			//$this->_sheet->removeRow($r);
			//$r--;
			return PHPTemplate_Util::ROW_TYPE_END;
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_COMMENT){
			//$this->_sheet->removeRow($r);
			//$r--;
			return $type;
		}
		elseif ($type==PHPTemplate_Util::ROW_TYPE_ELSEIF
				|| $type==PHPTemplate_Util::ROW_TYPE_ELSE){
			if(get_class($objRows) == "PHPTemplate_If"){
				$objRows->addCondition($objCell->getText(), $this->_sheet);
				//$this->_sheet->removeRow($r);
				//$r--;
				return $type;
			}
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_VAR){
			//パラメータ取得　$this->_paramArray
			$this->addParamFromText($objCell->getText());
			//$this->_sheet->removeRow($r);
			//$r--;
			return $type;
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_EXEC
			|| $type == PHPTemplate_Util::ROW_TYPE_RESUME){
			$controlObj = $this->getControlObj($type, $objCell, $this->_sheet);
			//$objCellからExec情報をセット
			$objRows->addRow($controlObj);
			//$this->_sheet->removeRow($r);
			//$r--;
			return $type;
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_FONT){
			$this->_page->addFont(new PHPTemplate_Font($objCell->getText()));
			//$this->_sheet->removeRow($r);
			//$r--;
			return $type;
		}
		else{
			//制御文なら#endまで
			if($type == PHPTemplate_Util::ROW_TYPE_FOREACH
			|| $type == PHPTemplate_Util::ROW_TYPE_HFOREACH
			|| $type == PHPTemplate_Util::ROW_TYPE_WHILE
			|| $type == PHPTemplate_Util::ROW_TYPE_IF
			|| $type == PHPTemplate_Util::ROW_TYPE_PAGEHEADER
			|| $type == PHPTemplate_Util::ROW_TYPE_PAGEFOOTER){
				$controlObj = $this->getControlObj($type, $objCell, $this->_sheet);

				$r++;
				//$this->_sheet->removeRow($r);
				//次行から#endまで
				$endR=FALSE;
				for(/*$r=$r;*/; $r<=$rowMax; $r++){
				//for($r=$r; $r<=$this->_sheet->getHighestRow(); /*$r++*/){
					$tes = $this->_sheet->getHighestRow();
					if($r >=  $tes){
						$endR=TRUE;
					}

					$rowObj = new PHPTemplate_Row();
					for($c=0; $c<=$colMax; $c++) {
						$cell = $this->_sheet->getCellByColumnAndRow($c, $r);
						$objCell = new PHPTemplate_Cell($cell, -1, -1);
						$type = $this->getBlock($objCell, $r, $rowObj, $colMax, $rowMax, $controlObj);
						if($type == PHPTemplate_Util::ROW_TYPE_END){
							//#rowArray にobjをセット
							$objRows->addRow($controlObj);
							$className = get_class($controlObj);
							if($className=="PHPTemplate_PageHeader"){
								if($this->_page->getHeader()==NULL){
									$this->_page->setHeader($controlObj);
								}
							}
							elseif ($className=="PHPTemplate_PageFooter"){
								if($this->_page->getFooter()==NULL){
									$this->_page->setFooter($controlObj);
								}
							}
							return PHPTemplate_Util::ROW_TYPE_BLOCK;
						}
						elseif ($type == PHPTemplate_Util::ROW_TYPE_NONE){
							// 次のセルも取得
						}
						elseif ($type == PHPTemplate_Util::ROW_TYPE_BLOCK
								|| $type == PHPTemplate_Util::ROW_TYPE_ELSEIF
								|| $type == PHPTemplate_Util::ROW_TYPE_ELSE
								|| $type == PHPTemplate_Util::ROW_TYPE_VAR
								|| $type == PHPTemplate_Util::ROW_TYPE_EXEC
								|| $type == PHPTemplate_Util::ROW_TYPE_PAGEBREAK
								|| $type == PHPTemplate_Util::ROW_TYPE_COMMENT
								|| $type == PHPTemplate_Util::ROW_TYPE_RESUME
								|| $type == PHPTemplate_Util::ROW_TYPE_FONT){
							$rowObj = NULL;
							break;
						}
						else{
							// 制御文
							break;
						}
					}
					//$objに行追加
					if($rowObj != NULL){
						$controlObj->addRow($rowObj);
						//$this->_sheet->removeRow($r);
						//$r--;
					}
					if($endR){
						break;
					}
				}
				//#endなし
				throw new PHPTemplate_Exception('#end not found');
			}
		}
	}

	/**
	 *#var 宣言文からパラメータを取得
	 *
	 *@param	string		$text
	* */
	private function addParamFromText($text)
	{
		$text = trim($text);
		$pos = strpos($text, "#var");
		//#varを削除
		if($pos!== false){
			//OK
			$text = trim(substr($text, strlen("#var")));
			$tok = strtok($text, ",");
			if($tok === false){
				$this->splitParam($text);
			}
			else{
				while ($tok !== false) {
					$this->splitParam($tok);
					$tok = strtok(",");
				}
			}
		}
	}

	/**
	 *name=valueを分割し、_paramArrayにセット
	 *
	 *@param		string		$text
	* */
	private function splitParam($text)
	{
		$pos = strpos($text, "=");
		if($pos === false){
			$this->_paramArray[$text]=NULL;
		}
		else{
			$name = trim(substr($text, 0, $pos));
			$val = trim(substr($text, $pos+1));
			//文字列と数値を判定する
			if(substr($val, 0, 1)=="'" && substr($val, strlen($val)-1)=="'"){
				$this->_paramArray[$name] = trim($val, "'");
			}
			elseif (substr($val, 0, 1)=="\"" && substr($val, strlen($val)-1)=="\""){
				$this->_paramArray[$name] = trim($val, "\"");
			}
			else{
				if(strpos($val, ".") !== false){
					$this->_paramArray[$name] = floatval($val);
				}
				else{
					$this->_paramArray[$name] = intval($val);
				}
			}
		}
	}

	/**
	 *
	 *
	 *@param		int		$index
	 * */
	//private function getRowDimensions($index)
	//{
		//return $this->_sheet->getRowDimension($index,FALSE);
		/*$this->_sheet = $this->_sheet->refreshRowDimensions();
		$this->_rowDimensions = $this->_sheet->getRowDimensions();
		$retDimension=NULL;
		foreach($this->_rowDimensions as $dimension){
			if($dimension->getRowIndex() == $index){
				//print_r($dimension);
				$retDimension = $dimension;
			}
		}
		return $retDimension;*/
	//}


	/**
	 *制御オブジェクトの生成
	 *
	 *@param	int					$type
	 *@param	PHPTemplate_Cell	$objCell
	 *@param	PHPExcel_Sheet		$sheet]
	 *@return	mixed
	* */
	static public function getControlObj($type, $objCell, $sheet)
	{
		$objCellText = $objCell->getText();
		if($type == PHPTemplate_Util::ROW_TYPE_FOREACH){
			return new PHPTemplate_Foreach($objCellText, $sheet);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_HFOREACH){
			return new PHPTemplate_Hforeach($objCellText, $sheet, $objCell->getColumn());
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_WHILE){
			return new PHPTemplate_While($objCellText, $sheet);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_IF){
			return new PHPTemplate_If($objCellText, $sheet);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_EXEC){
			return new PHPTemplate_Exec($objCellText);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_RESUME){
			return new PHPTemplate_Resume($objCellText, $sheet);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_PAGEHEADER){
			return new PHPTemplate_PageHeader($objCellText, $sheet);
		}
		elseif ($type == PHPTemplate_Util::ROW_TYPE_PAGEFOOTER){
			return new PHPTemplate_PageFooter($objCellText, $sheet);
		}

	}

	static public function replaceValues($sheet, $values) {

		$rowMax = $sheet->getHighestRow();
		$colMax = PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn());
		for($r=1; $r<=$rowMax; $r++) {
			for($c=0; $c<=$colMax; $c++) {
				$cell = $sheet->getCellByColumnAndRow($c, $r);

				$text = $cell->getValue();
				if (is_object($text)) {

					$rtfCell = $text->getRichTextElements();
					foreach ($rtfCell as $v) {
						foreach ($values as $val){
							$v->setText(str_replace($val->getSearchVal(), $val->getReplaceVal(), $v->getText()));
						}
					}
					$text->setRichTextElements($rtfCell);
				}
				else{
					foreach ($values as $val){
						//echo "getSearchVal:".$val->getSearchVal()."\n";
						//echo "getReplaceVal:".$val->getReplaceVal()."\n";
						$text = str_replace($val->getSearchVal(), $val->getReplaceVal(), $text);
					}
				}
				//echo "text:".$text."\n";
				$cell->setValue($text);
			}
		}
		return $sheet;
	}

}