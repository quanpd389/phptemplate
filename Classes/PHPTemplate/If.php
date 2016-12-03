<?php
/**
 * PHPTemplate_IfCondition
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_IfCondition
 * @copyright
 */
class PHPTemplate_IfCondition extends PHPTemplate_ControlBase
{
	/**
	 *  条件文
	 *
	 *  @var  string
	 */
	public $_condition = NULL;

	/**
	 *	Create a new IFCondition
	 *
	 *	@param	string					$text　
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$this->_sheet = $sheet;
		$text = trim($text);
		$pos = strpos($text, "#if");
		//#IFを削除
		if($pos!== false){
			//OK
			$text = trim(substr($text, strlen("#if")));
			$this->_condition = $text;
		}
		else{
			$pos = strpos($text, "#else if");
			if($pos !== false){
				//OK
				$text = trim(substr($text, strlen("#else if")));
				$this->_condition = $text;
			}
			else{
				$this->_condition = NULL;
			}
		}

	}

}

/**
 * PHPTemplate_If
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_If
 * @copyright
 */
class PHPTemplate_If
{
	/**
	 *  条件
	 *
	 *  @var  PHPTemplate_IfCondition[]
	 */
	private $_conditionArray = Array();

	/**
	 *  実行行
	 *
	 *  @var  mixed
	 */
	private $_current = NULL;

	/**
	 *	Create a new IF
	 *
	 *	@param	string					$text　
	 *	@param	PHPExcel_Worksheet		$sheet
	 */
	public function __construct($text, $sheet)
	{
		$conObj = new PHPTemplate_IfCondition($text, $sheet);
		$this->_conditionArray[] = $conObj;
		$this->_current = $conObj;

	}

	/**
	 * 行追加
	 *
	 * @param	PHPTemplate_Row			$obj
	 */
	public function addRow($obj){
		$this->_current->addRow($obj);
	}

	/**
	 * 条件追加
	 *
	 * @param	string					$text
	 * @param	PHPExcel_Worksheet		$sheet
	 */
	public function addCondition($text, $sheet){
		// PHPTemplate_IfCondition
		$conObj = new PHPTemplate_IfCondition($text, $sheet);
		$this->_conditionArray[] = $conObj;
		$this->_current = $conObj;
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
		foreach($this->_conditionArray as $con){
			$ret = PHPTemplate_Util::evalCondition($con->_condition, $data, $rownum, $page->getPagenum());
			if($ret){
				//行実行
				//$con->getRowArray
				$newRowArray = Array();
				$startRow = $rownum;
				for($i=0; $i < $con->getRowArrayCount(); $i++){
					$obj = $con->getRowArray($i);
					$className = get_class($obj);
					if($className == "PHPTemplate_Row"){
						//行追加
						if($dataF){
							$newRow = $obj->addRow($rownum);
						}else{
							$newRow = $obj->addRow($rownum, FALSE);
						}
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
						}
					}
				}*/
				break;
			}
		}
	}

}