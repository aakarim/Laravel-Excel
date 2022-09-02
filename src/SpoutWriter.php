<?php

namespace Maatwebsite\Excel;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Maatwebsite\Excel\Contracts\Sheet;
use Maatwebsite\Excel\Contracts\Writer;
use Maatwebsite\Excel\Events\BeforeWriting;
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
        $reader = ReaderEntityFactory::createReaderFromFile($path);
        $reader->open($path);

        $ss = [];
        // read everything into memory 
        foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
            $arrayIndex = $sheetIndex - 1;
            foreach ($sheet->getRowIterator() as $row) {
                if (is_null($ss[$arrayIndex] ?? null)) {
                    $ss[$arrayIndex] = [];
                }
                array_push($ss[$arrayIndex], $row);
            }
        }
        $reader->close();

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
        
        $writer = WriterEntityFactory::createWriterFromFile($path);
        $writer->openToFile($path);
        foreach ($this->spreadsheet->spreadsheet as $sheetIndex => $sheet) {
            if ($sheetIndex !== 0) {
                $writer->addNewSheetAndMakeItCurrent();
            }

            foreach ($sheet as $row) {
                $writer->addRow($row);
            }
        }
        $writer->close();
        
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
