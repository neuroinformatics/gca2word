<?php

namespace NIJC;

/**
 * image converter class.
 */
class ImageConverter
{
    /**
     * convert to white background png image.
     *
     * @param string $imagePath
     * @param string $outputPath
     *
     * @return string
     */
    public static function convertToWhiteBackgroundPNG($imagePath, $outputPath)
    {
        $size = getimagesize($imagePath);
        if (false === $size) {
            return false;
        }
        list($width, $height, $type) = $size;
        switch ($type) {
        case IMAGETYPE_GIF:
            $imSrc = @imagecreatefromgif($imagePath);
            break;
        case IMAGETYPE_JPEG:
            $imSrc = @imagecreatefromjpeg($imagePath);
            break;
        case IMAGETYPE_PNG:
            $imSrc = @imagecreatefrompng($imagePath);
            break;
        case IMAGETYPE_WBMP:
            $imSrc = @imagecreatefromwbmp($imagePath);
            break;
        case IMAGETYPE_XPM:
            $imSrc = @imagecreatefromxpm($imagePath);
            break;
        //case IMAGETYPE_WEBP:
        //    $im = @imagecreatefromwebp($imagePath);
        //    break;
        //case IMAGETYPE_BMP:
        //    $im = @imagecreatefrombmp($imagePath);
        //    break;
        default:
            return false;
        }
        if (false === $imSrc) {
            return false;
        }
        imageinterlace($imSrc, true);
        $imDest = imagecreatetruecolor($width, $height);
        imageinterlace($imDest, true);
        $bgColor = imagecolorallocate($imDest, 255, 255, 255);
        imagefilledrectangle($imDest, 0, 0, $width, $height, $bgColor);
        imagecopy($imDest, $imSrc, 0, 0, 0, 0, $width, $height);
        $ret = imagepng($imDest, $outputPath);
        imagedestroy($imSrc);
        imagedestroy($imDest);

        return $ret;
    }
}
