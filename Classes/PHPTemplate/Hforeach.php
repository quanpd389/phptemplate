<?php
/**
 * PHPTemplate_Hforeach
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Hforeach
 * @copyright
 */
class PHPTemplate_Hforeach extends PHPTemplate_ControlBase
{
	/**
	 *  バインドデータ名
	 *
	 *  @var  string
	 */
	private $_list;
	/**
	 *  データ参照名
	 *
	 *  @var  string
	 */
	private $_var;
	/**
	 * インデックス名
	 *
	 *  @var  string
	 */
	private $_indexname = "index";
	/**
	 * ループ開始カラム
	 *
	 *  @var  string
	 */
	private $_topcolumn = "A";
	/**
	 * ループカラム数
	 *
	 *  @var  int
	 */
	private $_columnCnt = 0;
	/**
	 * ヘッダ部分のカラム数
	 *
	 *  @var  int
	 */
	private $_headColumn = 0;


	/**
	 *	Create a new Hforeach
	 *
	 *	@param	string					$text	hforeach parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 **	@param	string 					$column
	 */
	public function __construct($text, $sheet, $column)
	{
		$this->_sheet = $sheet;
		$this->_topcolumn = $column;
		$text = trim($text);
		$pos = strpos($text, " ");
		//#hforeachを削除
		if($pos!== false && substr($text, 0, $pos) == "#hforeach"){
			//OK,
			$text = trim(substr($text, $pos));
			$pos = strpos($text, ":");
			if($pos !== false){
				//list取得
				$this->_var = trim(substr($text, 0, $pos));
				$text = trim(substr($text, $pos+1));
				//var取得
				$pos = strpos($text, " ");
				if($pos !== false){
					$this->_list =  trim(substr($text, 0, $pos));
					$text = trim(substr($text, $pos));
				}
				else{
					if(strlen($text) > 0){
						$this->_list = $text;
						return;
					}
					else{
						//NG
						throw new PHPTemplate_Exception('#hforeach parameter is invalid');
					}
				}
			}
			else{
				//NG
				throw new PHPTemplate_Exception('#hforeach parameter is invalid');
			}
		}
		else{
			//NG
			throw new PHPTemplate_Exception('#hforeach parameter is invalid');
		}
		//index取得
		$p = PHPTemplate_Util::getParamater($text, "index");
		if($p != NULL)	$this->_indexname = $p;

	}

	public function setTopColumn($column){
		$this->_topcolumn = $column;

	}

	public function addColumn($column){
		$this->_columnArray[] = $column;
	}


	/**
	 * 制御ブロックの展開
	 *
	 * @param array				$data
	 * @param int				$rownum
	 * @param PHPTemplate_Page	$page
	 * @param bool				$dataF
	 */
	public function expand(&$data, &$rownum, &$page, $dataF=TRUE)
	{
		//バインド対象のカラムを取得
		for($i=0; $i < $this->getRowArrayCount(); $i++){
			$obj = $this->getRowArray($i);
			$className = get_class($obj);
			$f = FALSE;
			$cnt=0;

			if($className == "PHPTemplate_Row"){
				for($idx=0; $idx < $obj->getCellCnt() ;$idx++){
					$cell = $obj->getCell($idx);
					if($cell->getColumn() == $this->_topcolumn){
						$f = TRUE;
						$this->_headColumn=$idx;
					}
					if($f){
						if(strlen($cell->getText()) > 0){
							$cnt++;
							if($this->_columnCnt < ($cnt)){
								$this->_columnCnt = $cnt;
							}
						}
					}
				}
			}
		}
		//全体のカラム数は$headColumn+$this->_columnCnt
		for($i=0; $i < $this->getRowArrayCount(); $i++){
			$obj = $this->getRowArray($i);
			$className = get_class($obj);
			if($className == "PHPTemplate_Row"){
				$obj->trimRow($this->_headColumn + $this->_columnCnt);
			}
		}

		// データ取得
		$dataList = PHPTemplate_Util::getBindData($this->_list, $data, $rownum, $page->getPagenum());
		if($dataList == NULL){
			return;//error
		}
		// ヘッダー部分を展開
		for($i=0; $i < $this->getRowArrayCount(); $i++){
			$obj = $this->getRowArray($i);
			$className = get_class($obj);
			if($className == "PHPTemplate_Row"){
				$mergeSt = NULL;
				$mergeEd = NULL;
				for($j=0; $j < $this->_headColumn+$this->_columnCnt; $j++){
					$cell = $obj->getCell($j);
					$this->_sheet->duplicateStyle($cell->getStyle(), $cell->getColumn().($rownum+$i));
					$this->_sheet->getColumnDimension($cell->getColumn())->setWidth($cell->getColumnWidth());
					$this->_sheet->setCellValue($cell->getColumn().($rownum+$i),$cell->getText());
					if($j < $this->_headColumn)
						$cell->bindData($data, $rownum+$i, $page);
					//結合情報を取得
					$info = $cell->getMergeInfo();
					if($info[0]==$cell->getColumn()){
						$mergeSt = $cell->getColumn();
					}
					elseif ($info[1]==$cell->getColumn()){
						$mergeEd = $cell->getColumn();
						if($mergeSt!=NULL && $mergeEd!=NULL){
							$merge = $mergeSt.(string)($rownum+$i).":".$mergeEd.(string)($rownum+$i);
							$this->_sheet->mergeCells($merge);
						}
					}
				}
			}
		}
		//ヘッダー部分行結合
		for($i=0; $i < $this->getRowArrayCount(); $i++){
			$obj = $this->getRowArray($i);
			$className = get_class($obj);
			if($className == "PHPTemplate_Row"){
				foreach($obj->getRowMergeInfo() as $info){
					if($info->getRowCnt() > 0){
						$mi = $info->getMergeInfo();
						$merge = $mi[0].(string)($rownum+$i).":".$mi[1].(string)($rownum+$i+$info->getRowCnt());
						$this->_sheet->mergeCells($merge);
					}
				}
			}
		}

		// カラムを追加
		for($k=0; $k < (count($dataList)-1); $k++){
			$newColumnArray = Array();
			for($i=0; $i < $this->getRowArrayCount(); $i++){
				$obj = $this->getRowArray($i);
				$className = get_class($obj);
				if($className == "PHPTemplate_Row"){
					$sheet = NULL;
					$mergeSt = NULL;
					$mergeEd = NULL;
					for($j=$this->_headColumn; $j < $this->_headColumn+$this->_columnCnt; $j++){
						$cell = $obj->getCell($j);
						$sheet = $cell->getWorksheet();
						$style = $cell->getStyle();

						$c = PHPExcel_Cell::stringFromColumnIndex($j+($this->_columnCnt*($k+1)));

						$sheet->duplicateStyle($style, $c.($rownum+$i));
						$sheet->getColumnDimension($c)->setWidth($cell->getColumnWidth());
						$sheet->setCellValue($c.($rownum+$i),$cell->getText());
						$objCell = $sheet->getCell($c.($rownum+$i));
						$newCell = new PHPTemplate_Cell($objCell, $cell->getMergeRowHeight(), $cell->getMergeColumnWidth());
						$obj->addCell($newCell);
						//$idx = PHPExcel_Cell::stringFromColumnIndex($j-($this->_columnCnt*($k+1)));
						$newColumnArray[$cell->getColumn()] = $c;

						//結合情報を取得
						$info = $cell->getMergeInfo();
						if($info[0]==$cell->getColumn()){
							$mergeSt = $c;
						}
						elseif ($info[1]==$cell->getColumn()){
							$mergeEd = $c;
							if($mergeSt!=NULL && $mergeEd!=NULL){
								$merge = $mergeSt.(string)($rownum+$i).":".$mergeEd.(string)($rownum+$i);
								$sheet->mergeCells($merge);
							}
						}
					}
				}
			}
			//行結合
			for($i=0; $i < $this->getRowArrayCount(); $i++){
				$obj = $this->getRowArray($i);
				$className = get_class($obj);
				if($className == "PHPTemplate_Row"){
					foreach($obj->getRowMergeInfo() as $info){
						if($info->getRowCnt() > 0){
							$mi = $info->getMergeInfo();
							if(array_key_exists($mi[0], $newColumnArray)&&array_key_exists($mi[1], $newColumnArray)){
								$merge = $newColumnArray[$mi[0]].(string)($rownum+$i).":".$newColumnArray[$mi[1]].(string)($rownum+$i+$info->getRowCnt());
								$this->_sheet->mergeCells($merge);
							}
						}
					}
				}
			}
		}
		// 追加カラムにバインド
		$count = 0;
		foreach ($dataList as $d){
			$data = $data;
			$data[$this->_var] = $d;
			$data[$this->_indexname] = $count;

			$top = $this->_headColumn + ($this->_columnCnt*$count);

			for($i=0; $i < $this->getRowArrayCount(); $i++){
				$obj = $this->getRowArray($i);
				$className = get_class($obj);
				if($className == "PHPTemplate_Row"){
					for($j=$top; $j < ($this->_headColumn+($this->_columnCnt*($count+1))); $j++){
						$cell = $obj->getCell($j);
						$cell->bindData($data, $rownum+$i, $page);
					}
				}
			}
			$count++;
		}
		$rownum = $rownum + $this->getRowArrayCount();
	}

}