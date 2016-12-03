<?php
/**
 * PHPTemplate_While
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_While
 * @copyright
 */
class PHPTemplate_While extends PHPTemplate_ControlBase
{
	/**
	 *  whileの条件文
	 *
	 *  @var  string
	 */
	private $_script = NULL;


	/**
	 *	Create a new While
	 *
	 *	@param	string					$text	parameter
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;
		$pos = strpos($text, "#while");
		if($pos !== false){
			$this->_script = trim(substr($text, strlen("#while")));
		}
		else{
			throw new PHPTemplate_Exception('#while parameter is invalid');
		}

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
		while(PHPTemplate_Util::evalCondition($this->_script, $data, $rownum, $page->getPagenum())){
			//行実行
			$newRowArray = Array();
			$startRow = $rownum;
			for($i=0; $i<$this->getRowArrayCount(); $i++){
				$obj = $this->getRowArray($i);
				$className = get_class($obj);

				if($className == "PHPTemplate_Row"){
					//行追加
					$newRow = $obj->addRow($rownum, $dataF);
					//データバインド
					$newRow->expand($data, $rownum, $page);
					$newRowArray[] = $newRow;
				}
				else{
					$obj->expand($data, $rownum, $page, $dataF);
				}
			}
			/*foreach ($newRowArray as $row){
				foreach($row->getRowMergeInfo() as $info){
					if($info->getRowCnt() > 0){
						$p = $info->getMergeInfo();
						$merge = $p[0].(string)($startRow).":".$p[1].(string)($startRow+$info->getRowCnt());
						$this->_sheet->mergeCells($merge);
						//echo "while".$merge;
					}
				}
				$startRow++;
			}*/
		}

	}

}
