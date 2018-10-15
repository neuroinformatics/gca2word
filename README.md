# gca2word

Create MS-Word abstracts book from GCA-Web online abstracts

## prepare
- install [texvc](https://www.mediawiki.org/wiki/Texvc) command.
- run 'composer install' command.
- write metadata file CONFERENCE.json and put to config/conference/
- modify default connection settings config/default.json if needed.
  (it can be override by command line parameter)

## usage
```
Usage: gca2ward.php [options]

Options:
  -h HOSTNAME           url of GCA-Web
                        (default "https://abstracts.neuroinf.jp/")
  -u USERNAME           login id (mail address) of GCA-Web
                        (default "admin@ni.riken.jp")
  -p                    input password from command line
  -c CONFERENCE         target conference short name (default "AINI2017")
  -o OUTPUTFILE         output file name (default "abstract.docx")
