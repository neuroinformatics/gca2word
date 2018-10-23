<?php

namespace NIJC;

/**
 * password utility class.
 */
class PasswordUtility
{
    /**
     * input password from tty.
     *
     * @param int $minLength
     * @param int $maxLength
     *
     * @return string
     */
    public static function input($minLength = 0, $maxLength = 0)
    {
        while (1) {
            $pass = self::_ttyInput('Enter Password:');
            $len = strlen($pass);
            if ($len >= $minLength && (0 == $maxLength || $len <= $maxLength)) {
                break;
            }
            echo sprintf('You must type in %u to %u characters.', $minLength, $maxLength).PHP_EOL;
        }

        return $pass;
    }

    /**
     * input password with confirmation from tty.
     *
     * @param int $minLength
     * @param int $maxLength
     *
     * @return string
     */
    public static function inputWithConfirm($minLength = 0, $maxLength = 0)
    {
        $pass = self::input($minLength, $maxLength);
        for ($try = 0; 1; ++$try) {
            if ($try > 2) {
                return null;
            }
            $pass2 = self::_ttyInput('Enter Password (Confirm):');
            if ($pass == $pass2) {
                break;
            }
            echo 'Password mismatch'.PHP_EOL;
        }

        return $pass;
    }

    /**
     * get password from tty.
     *
     * @param string $message
     *
     * @return string
     */
    protected static function _ttyInput($message)
    {
        fwrite(STDERR, $message);
        system('stty -echo');
        $ret = trim(fgets(STDIN));
        system('stty echo');
        echo PHP_EOL;

        return $ret;
    }
}
