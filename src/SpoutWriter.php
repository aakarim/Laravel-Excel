<?php

namespace Maatwebsite\Excel;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Maatwebsite\Excel\Contracts\Sheet;
use Maatwebsite\Excel\Contracts\Writer;
use Maatwebsite\Excel\Events\BeforeWriting;
use Maatwebsite\Excel\Files\LocalTemporaryFile;
use Maatwebsite\Excel\Files\RemoteTemporaryFile;
use Maatwebsite\Excel\Files\TemporaryFile;
use Maatwebsite\Excel\Files\TemporaryFileFactory;

class SpoutWriter implements Writer
{
    use HasEventBus;

    protected SpoutBook $spreadsheet;

    /**
     * @var object
     */
    protected $exportable;

    /**
     * @var TemporaryFileFactory
     */
    protected $temporaryFileFactory;

    /**
     * @param TemporaryFileFactory $temporaryFileFactory
     */
    public function __construct(TemporaryFileFactory $temporaryFileFactory)
    {
        $this->temporaryFileFactory = $temporaryFileFactory;
    }

    /**
     * @param TemporaryFile $tempFile
     * @param string        $writerType
     *
     * @return Writer
     */
    public function reopen(TemporaryFile $tempFile, string $writerType): SpoutWriter
    {
        $path = $tempFile->sync()->getLocalPath();
 
        $ss = [];

        $this->spreadsheet = new SpoutBook($ss);
        return $this;
    }

    /**
     * @param object        $export
     * @param TemporaryFile $temporaryFile
     * @param string        $writerType
     *
     * @return TemporaryFile
     */
    public function write($export, TemporaryFile $temporaryFile, string $writerType): TemporaryFile
    {
        $this->exportable = $export;
        
        $this->raise(new BeforeWriting($this, $this->exportable));

        $path = $temporaryFile->sync()->getLocalPath();
        $pathParts = pathinfo($path);

        $reader = ReaderEntityFactory::createReaderFromFile($path);
        $reader->open($path);
        
        // create new temporary file for the witer 
        // TODO: create a remote file as appropriate
        $writerFilePath = $pathParts['dirname'] . '/' . $pathParts['filename'] . '_writer' . '.' . $pathParts['extension'];
        $tmpWriterFile = new LocalTemporaryFile($writerFilePath);
        $writer = WriterEntityFactory::createWriterFromFile($tmpWriterFile->sync()->getLocalPath());
        $writer->openToFile($tmpWriterFile->sync()->getLocalPath());
        // copy the existing file over
        foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
            // Add sheets in the new file, as we read new sheets in the existing one
            if ($sheetIndex !== 1) {
                $writer->addNewSheetAndMakeItCurrent();
            }
            foreach ($sheet->getRowIterator() as $row) {
                // ... and copy each row into the new spreadsheet
                $writer->addRow($row);
            }
        }
        $reader->close();

        // then append rows using the addition in memory
        // assume that all the sheets have been created already
        foreach ($this->spreadsheet->spreadsheet as $sheetIndex => $sheet) {
            $writerSheet = $writer->getSheets()[$sheetIndex];
            $writer->setCurrentSheet($writerSheet);
            
            foreach ($sheet as $row) {
                $writer->addRow($row);
            }
        }
        $writer->close();
        
        // TODO: change when we support local and remote files
        unlink($path);
        rename($writerFilePath, $path);

        if ($temporaryFile instanceof RemoteTemporaryFile) {
            $temporaryFile->updateRemote();
            $temporaryFile->deleteLocalCopy();
        }

        $this->clearListeners();
        unset($this->spreadsheet);

        return $temporaryFile;
    }

    /**
     * @param int $sheetIndex
     *
     * @return Sheet
     */
    public function getSheetByIndex(int $sheetIndex): Sheet
    {
        return new SpoutSheet($this->spreadsheet, $sheetIndex);
    }
}
