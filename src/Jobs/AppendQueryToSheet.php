<?php

namespace Maatwebsite\Excel\Jobs;

use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Maatwebsite\Excel\Concerns\FromQuery;
use Maatwebsite\Excel\Files\TemporaryFile;
use Maatwebsite\Excel\Jobs\Middleware\LocalizeJob;
use Maatwebsite\Excel\Writer;

class AppendQueryToSheet implements ShouldQueue
{
    use Queueable, Dispatchable, ProxyFailures, InteractsWithQueue;

    /**
     * @var TemporaryFile
     */
    public $temporaryFile;

    /**
     * @var string
     */
    public $writerType;

    /**
     * @var int
     */
    public $sheetIndex;

    /**
     * @var FromQuery
     */
    public $sheetExport;

    /**
     * @var int
     */
    public $page;

    /**
     * @var int
     */
    public $chunkSize;

    /**
     * @param  FromQuery  $sheetExport
     * @param  TemporaryFile  $temporaryFile
     * @param  string  $writerType
     * @param  int  $sheetIndex
     * @param  int  $page
     * @param  int  $chunkSize
     */
    public function __construct(
        FromQuery $sheetExport,
        TemporaryFile $temporaryFile,
        string $writerType,
        int $sheetIndex,
        int $page,
        int $chunkSize
    ) {
        $this->sheetExport   = $sheetExport;
        $this->temporaryFile = $temporaryFile;
        $this->writerType    = $writerType;
        $this->sheetIndex    = $sheetIndex;
        $this->page          = $page;
        $this->chunkSize     = $chunkSize;
    }

    /**
     * Get the middleware the job should be dispatched through.
     *
     * @return array
     */
    public function middleware()
    {
        return (method_exists($this->sheetExport, 'middleware')) ? $this->sheetExport->middleware() : [];
    }

    /**
     * @param  Writer  $writer
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function handle(Writer $writer)
    {
        (new LocalizeJob($this->sheetExport))->handle($this, function () use ($writer) {
            $parent = \Sentry\SentrySdk::getCurrentHub()->getSpan();
            $span = null;

            // Check if we have a parent span (this is the case if we started a transaction earlier)
            if ($parent !== null) {
                $context = new \Sentry\Tracing\SpanContext();
                $context->setOp('exports.queue.reopen');
                $context->setDescription('Reopening temporary file');
                $span = $parent->startChild($context);

                // Set the current span to the span we just started
                \Sentry\SentrySdk::getCurrentHub()->setSpan($span);
            }
            
            $writer = $writer->reopen($this->temporaryFile, $this->writerType);
            
            if ($span !== null) {
                $span->finish();
                
                // Restore the current span back to the parent span
                \Sentry\SentrySdk::getCurrentHub()->setSpan($parent);
            }
            
            $sheet = $writer->getSheetByIndex($this->sheetIndex);
            
            if ($parent !== null) {
                $context = new \Sentry\Tracing\SpanContext();
                $context->setOp('exports.queue.query');
                $context->setDescription('executing db query for rows');
                $span = $parent->startChild($context);

                // Set the current span to the span we just started
                \Sentry\SentrySdk::getCurrentHub()->setSpan($span);
            }
            $query = $this->sheetExport->query()->forPage($this->page, $this->chunkSize);
            if ($span !== null) {
                $span->finish();
                
                // Restore the current span back to the parent span
                \Sentry\SentrySdk::getCurrentHub()->setSpan($parent);
            }
            
            if ($parent !== null) {
                $context = new \Sentry\Tracing\SpanContext();
                $context->setOp('exports.queue.appendRows');
                $context->setDescription('appending rows to sheet');
                $span = $parent->startChild($context);

                // Set the current span to the span we just started
                \Sentry\SentrySdk::getCurrentHub()->setSpan($span);
            }
            $sheet->appendRows($query->get(), $this->sheetExport);
            if ($span !== null) {
                $span->finish();
                
                // Restore the current span back to the parent span
                \Sentry\SentrySdk::getCurrentHub()->setSpan($parent);
            }
            
            
            if ($parent !== null) {
                $context = new \Sentry\Tracing\SpanContext();
                $context->setOp('exports.queue.write');
                $context->setDescription('writing to temporary file');
                $span = $parent->startChild($context);

                // Set the current span to the span we just started
                \Sentry\SentrySdk::getCurrentHub()->setSpan($span);
            }
            $writer->write($this->sheetExport, $this->temporaryFile, $this->writerType);
            if ($span !== null) {
                $span->finish();

                // Restore the current span back to the parent span
                \Sentry\SentrySdk::getCurrentHub()->setSpan($parent);
            }
        });
    }
}
