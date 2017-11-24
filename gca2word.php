#!/usr/bin/env php
<?php

require_once __DIR__.'/vendor/autoload.php';

$configJsonPath = __DIR__.'/config/default.json';

if (!file_exists($configJsonPath)) {
    die('Fatal Error: config json not found'.PHP_EOL);
}
$configJson = file_get_contents($configJsonPath);
$config = json_decode($configJson, true);
if (null === $config) {
    die('Fatal Error: invalid config json format'.PHP_EOL);
}

$options = getopt('h:u:pc:o:');
foreach ($options as $option => $value) {
    switch ($option) {
    case 'h':
        $config['host'] = trim($value);
        break;
    case 'u':
        $config['auth']['user'] = trim($value);
        break;
    case 'p':
        $value = NIJC\PasswordUtility::input();
        if (null === $value) {
            die('Fatal Error: failed to set password'.PHP_EOL);
        }
        $config['auth']['password'] = trim($value);
        break;
    case 'c':
        $config['conference'] = trim($value);
        break;
    case 'o':
        $config['output'] = trim($value);
        break;
    }
}

$gca = new NIJC\GcaClient($config['host'], $config['auth']['user'], $config['auth']['password']);
$conferenceJson = $gca->getConference($config['conference']);
if (false === $conferenceJson) {
    die('Fatal Error: failed to get conference information: '.$config['conference'].PHP_EOL);
}
$abstractsJson = $gca->getAbstracts($config['conference']);
if (false === $abstractsJson) {
    die('Fatal Error: failed to get all abstracts data: '.$config['conference'].PHP_EOL);
}
$tempDir = new NIJC\TempDirectory();
$imageDir = $tempDir->get();

if (!$gca->saveImages($abstractsJson, $imageDir)) {
    die('Fatal Error: failed to get image data: '.$config['conference'].PHP_EOL);
}

$writer = new NIJC\AbstractsWriter($abstractsJson, $conferenceJson, $imageDir);
$writer->save($config['output']);
