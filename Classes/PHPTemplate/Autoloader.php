<?php
PHPTemplate_Autoloader::Register();
if (ini_get('mbstring.func_overload') & 2) {
    throw new PHPTemplate_Exception('Multibyte function overloading in PHP must be disabled for string functions (2).');
}

class PHPTemplate_Autoloader
{
    /**
     * Register the Autoloader with SPL
     *
     */
    public static function Register() {
        if (function_exists('__autoload')) {
            //    Register any existing autoloader function with SPL, so we don't get any clashes
            spl_autoload_register('__autoload');
        }
        //    Register ourselves with SPL
        return spl_autoload_register(array('PHPTemplate_Autoloader', 'Load'));
    }   //    function Register()


    /**
     * Autoload a class identified by name
     *
     * @param    string    $pClassName        Name of the object to load
     */
    public static function Load($pClassName){
        if ((class_exists($pClassName,FALSE)) || (strpos($pClassName, 'PHPTemplate') !== 0)) {
            //    Either already loaded, or not a PHPTemplate class request
            return FALSE;
        }

        $pClassFilePath = PHPTEMPLATE_ROOT .
                          str_replace('_',DIRECTORY_SEPARATOR,$pClassName) .
                          '.php';

        if ((file_exists($pClassFilePath) === FALSE) || (is_readable($pClassFilePath) === FALSE)) {
            //    Can't load
            return FALSE;
        }

        require($pClassFilePath);
    }   //    function Load()

}
