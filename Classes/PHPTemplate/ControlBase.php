<?php
/**
 * PHPTemplate_ControlBase
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_ControlBase
 * @copyright
 */
class PHPTemplate_ControlBase
{
	/**
	 *  制御ブロック内の行リスト
	 *
	 *  @var  array
	 */
	private $_rowArray = array();

	/**
	 *  ワークシート
	 *
	 *  @var  PHPExcel_Worksheet
	 */
	protected  $_sheet = NULL;

	/**
	 * 行追加
	 *
	 * @param mixed $obj
	 */
	public function addRow($obj){
		// PHPTemplate_Row,PHPTemplate_Roreach etc...
		$this->_rowArray[] = $obj;
	}

	/**
	 * 行取得
	 * @param int $idx
	 * @return mixed
	 */
	public function getRowArray($idx){
		return $this->_rowArray[$idx];
	}
	/**
	 * 行数取得
	 *
	 * @return int
	 */
	public function getRowArrayCount(){
		return count($this->_rowArray);
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
		$newRowArray = Array();
		$startRow = $rownum;
		for($i=0; $i<$this->getRowArrayCount(); $i++){
			$obj = $this->getRowArray($i);
			$className = get_class($obj);

			if($className == "PHPTemplate_Row"){
				$newRow = $obj->addRow($rownum, $dataF);
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
					//echo "Base".$merge;
				}
			}
			$startRow++;
		}*/

	}



}