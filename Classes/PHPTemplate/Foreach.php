<?php
/**
 * PHPTemplate_Foreach
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Foreach
 * @copyright
 */
class PHPTemplate_Foreach extends PHPTemplate_ControlBase
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
	 * 最大指定行
	 *
	 *  @var  int
	 */
	private $_maxrowcount = 0;

	/**
	 *	Create a new Foreach
	 *
	 *	@param	string					$text	foreach parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;
		$text = trim($text);
		$pos = strpos($text, " ");
		//#foreachを削除
		if($pos!== false && substr($text, 0, $pos) == "#foreach"){
			//OK
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
						throw new PHPTemplate_Exception('#foreach parameter is invalid');
					}
				}
			}
			else{
				//NG
				throw new PHPTemplate_Exception('#foreach parameter is invalid');
			}
		}
		else{
			//NG
			throw new PHPTemplate_Exception('#foreach parameter is invalid');
		}
		//index, max 取得
		$p= PHPTemplate_Util::getParamater($text, "index");
		if($p != NULL)	$this->_indexname = $p;
		$p = PHPTemplate_Util::getParamater($text, "max");
		if($p != NULL)	$this->_maxrowcount = $p;
	}

	/**
	 * 制御ブロックの展開
	 *
	 * @param array				$data
	 * @param int				$rownum
	 * @param PHPTemplate_Page	$page
	 * @param bool				$dataF
	 */
	public function expand(&$data, &$rownum, &$page, $dataF=TRUE){
		// データ取得
		$dataList = PHPTemplate_Util::getBindData($this->_list, $data, $rownum, $page->getPagenum());
		if($dataList == NULL){
			return;//error
		}
		//maxが指定されている場合はページ単位でのMAXとする
		$pageListCnt=0;
		$currentPage=$page->getPagenum();
		for($count=0; ;$count++){
			if($this->_maxrowcount==0){
				if ($count >= count($dataList)){
					break;
				}
			}
			else{
				//max指定あり
				if($count >= count($dataList)){
					if($pageListCnt >= $this->_maxrowcount){
						break;
					}
				}

			}
			if($currentPage!=$page->getPagenum()){
				$pageListCnt=0;
				$currentPage=$page->getPagenum();
			}
			$d = NULL;
			if ($count < count($dataList)) {
				$d = $dataList[$count];
			}

			$data = $data;
			$data[$this->_var] = $d;
			$data[$this->_indexname] = $count;
			$startRow = $rownum;
			$newRowArray = Array();
			for($i=0; $i < $this->getRowArrayCount(); $i++){
				$obj = $this->getRowArray($i);
				$className = get_class($obj);
				if($className == "PHPTemplate_Row"){
					//行追加
					if($d!=NULL){
						$newRow = $obj->addRow($rownum);
					}else{
						$newRow = $obj->addRow($rownum, FALSE);
					}
					//データバインド
					$newRow->expand($data, $rownum, $page);
					$newRowArray[] = $newRow;
				}
				else{
					if($d!=NULL){
						$obj->expand($data, $rownum, $page);
					}
					else{
						$obj->expand($data, $rownum, $page, FALSE);
					}
				}
			}

			/*foreach ($newRowArray as $row){
				foreach($row->getRowMergeInfo() as $info){
					if($info->getRowCnt() > 0){
						$p =  $info->getMergeInfo();
						$merge = $p[0].(string)($startRow).":".$p[1].(string)($startRow+$info->getRowCnt());
						$this->_sheet->mergeCells($merge);
						//echo "RowMarge".$merge;
					}
				}
			}*/
			$pageListCnt++;
			$startRow++;
		}
	}


}