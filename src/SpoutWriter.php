<?php

namespace Maatwebsite\Excel;

use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\WriterInterface;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Maatwebsite\Excel\Cache\CacheManager;
use Maatwebsite\Excel\Concerns\WithCustomValueBinder;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;
use Maatwebsite\Excel\Concerns\WithProperties;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Contracts\Sheet;
use Maatwebsite\Excel\Contracts\Writer;
use Maatwebsite\Excel\Events\BeforeExport;
use Maatwebsite\Excel\Events\BeforeWriting;
use Maatwebsite\Excel\Factories\WriterFactory;
use Maatwebsite\Excel\Files\RemoteTemporaryFile;
use Maatwebsite\Excel\Files\TemporaryFile;
use Maatwebsite\Excel\Files\TemporaryFileFactory;

class SpoutWriter implements Writer
{
    use DelegatedMacroable, HasEventBus;

    protected array $spreadsheet;

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

        $this->setDefaultValueBinder();
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

        $ss = [];
        // read everything into memory 
        foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
            $ss[$sheetIndex] = [];
            
            foreach ($sheet->getRowIterator() as $row) {
                array_push($ss[$sheetIndex], $row);
            }
        }
        $reader->close();

        $this->spreadsheet = $ss;

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
        \Log::info("write using spout");
        
        $path = $temporaryFile->sync()->getLocalPath();
        
        $writer = WriterEntityFactory::createWriterFromFile($path);
        $writer->openToFile($path);

        foreach ($this->spreadsheet as $sheetIndex => $sheet) {
            if ($sheetIndex !== 1) {
                $writer->addNewSheetAndMakeItCurrent();
            }

            foreach ($sheet->getRowIterator() as $row) {
                // ... and copy each row into the new spreadsheet
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
     * @return Spreadsheet
     */
    public function getDelegate()
    {
        return $this->spreadsheet ;
    }

    /**
     * @param int $sheetIndex
     *
     * @return Sheet
     */
    public function getSheetByIndex(int $sheetIndex): Sheet
    {
        return new SpoutSheet($this->spreadsheet[$sheetIndex]);
    }
}
