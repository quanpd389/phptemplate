<?php
/**
 * PHPTemplate_Font
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Font
 * @copyright
 */
class PHPTemplate_Font
{
	/**
	 *  フォント情報格納タグ名
	 *
	 *  @var  string
	 */
	private $_tagname=NULL;

	/**
	 *  フォント color
	 *
	 *  @var  string
	 */
	private $_color='000000';
	/**
	 *  フォント bold
	 *
	 *  @var  bool
	 */
	private $_bold=NULL;
	/**
	 *  フォント italic
	 *
	 *  @var  bool
	 */
	private $_italic=NULL;
	/**
	 *  フォント underline
	 *
	 *  @var  bool
	 */
	private $_underline=NULL;


	/**
	 *	Create a new Font
	 *
	 *	@param	string $text	parameter
	 */
	public function __construct($text)
	{
		$pos = strpos($text, "#font");
		if($pos !== false){
			$str=trim(substr($text, strlen("#font")));

			$this->_tagname = PHPTemplate_Util::getParamater($str, "tag");
			$this->_color=PHPTemplate_Util::getParamater($str, "color");
			$this->_bold=PHPTemplate_Util::getParamater($str, "bold");
			$this->_italic=PHPTemplate_Util::getParamater($str, "italic");
			$this->_underline=PHPTemplate_Util::getParamater($str, "underline");
		}
		else{
			throw new PHPTemplate_Exception('#font parameter is invalid');
		}
	}

	/**
	 *	Get tag name
	 *
	 *	@return		string
	 */
	public function getTagname()
	{
		return $this->_tagname;
	}
	/**
	 *	Get font color
	 *
	 *	@return		string
	 */
	public function getColor()
	{
		return $this->_color;
	}
	/**
	 *	Get font bold
	 *
	 *	@return		bool
	 */
	public function getBold()
	{
		return $this->_bold;
	}
	/**
	 *	Get font italic
	 *
	 *	@return		bool
	 */
	public function getItalic()
	{
		return $this->_italic;
	}
	/**
	 *	Get font underline
	 *
	 *	@return		bool
	 */
	public function getUnderline()
	{
		return $this->_underline;
	}

}