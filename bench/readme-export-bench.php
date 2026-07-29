<?php

/*
 * README export benchmark.
 *
 * Reproduces the scenario shown in the README's "Benchmarks" section: an XLSX
 * export of N rows x C columns of random data, measured for both ways of
 * feeding FastExcel:
 *
 *   - "collection": the whole dataset is held in memory (a Collection), then
 *     exported. Simple, but peak memory grows with the dataset.
 *   - "generator":  rows are yielded one at a time, so OpenSpout streams them
 *     straight to disk and peak memory stays flat regardless of row count.
 *
 * Each mode is measured in its OWN php process — PHP does not return freed heap
 * to the OS, so running both in one process would inflate the second mode's
 * peak. The parent run spawns one child per mode to keep the numbers honest.
 *
 * Usage:
 *   php bench/readme-export-bench.php [rows] [cols] [iterations]
 *   php bench/readme-export-bench.php 10000 20 10
 *   php bench/readme-export-bench.php 10000 20 10 generator   # single mode (worker)
 *
 * Reports the average wall-clock time and peak process memory
 * (memory_get_peak_usage(true)) across the iterations. Time is noisy; the
 * memory numbers are the interesting part. Don't trust benchmarks.
 */

require getenv('BENCH_AUTOLOAD') ?: __DIR__.'/../vendor/autoload.php';

use Rap2hpoutre\FastExcel\FastExcel;

$rows = max(1, (int) ($argv[1] ?? 10000));
$cols = max(1, (int) ($argv[2] ?? 20));
$iterations = max(1, (int) ($argv[3] ?? 10));
$mode = $argv[4] ?? null;

// Parent: spawn one isolated child process per mode, then print both results.
if ($mode === null) {
    printf('XLSX export — %d rows x %d cols, average of %d runs (PHP %s)%s', $rows, $cols, $iterations, PHP_VERSION, PHP_EOL);
    foreach (['collection', 'generator'] as $m) {
        $cmd = sprintf(
            '%s %s %d %d %d %s',
            escapeshellarg(PHP_BINARY),
            escapeshellarg(__FILE__),
            $rows,
            $cols,
            $iterations,
            $m
        );
        echo rtrim((string) shell_exec($cmd)), PHP_EOL;
    }
    exit(0);
}

// Worker: measure a single mode in this fresh process.
$dir = sys_get_temp_dir().'/fastexcel-readme-bench-'.getmypid();
mkdir($dir);

// One random row of $cols columns: a mix of ints and strings, keyed col_0..col_N.
$makeRow = static function () use ($cols) {
    $row = [];
    for ($c = 0; $c < $cols; $c++) {
        $row['col_'.$c] = ($c % 3 === 0) ? mt_rand(1, 1_000_000) : 'val_'.mt_rand(1, 1_000_000);
    }

    return $row;
};

$callback = match ($mode) {
    // Everything materialised in memory up front.
    'collection' => function () use ($rows, $makeRow, $dir) {
        $data = [];
        for ($r = 0; $r < $rows; $r++) {
            $data[] = $makeRow();
        }
        (new FastExcel(collect($data)))->export($dir.'/collection.xlsx');
    },
    // Streamed row by row through a generator.
    'generator' => function () use ($rows, $makeRow, $dir) {
        $gen = (function () use ($rows, $makeRow) {
            for ($r = 0; $r < $rows; $r++) {
                yield $makeRow();
            }
        })();
        (new FastExcel($gen))->export($dir.'/generator.xlsx');
    },
    default => throw new InvalidArgumentException("Unknown mode: {$mode}"),
};

$result = bench($iterations, $callback);

array_map('unlink', glob($dir.'/*'));
rmdir($dir);

printf('  FastExcel (%-10s) %7.3f s  %8.1f MB peak', $mode, $result['avg_s'], $result['peak_mb']);

function bench(int $iterations, callable $callback): array
{
    $times = [];
    $peak = 0;

    for ($i = 0; $i < $iterations; $i++) {
        gc_collect_cycles();
        if (function_exists('memory_reset_peak_usage')) {
            memory_reset_peak_usage();
        }
        $start = hrtime(true);
        $callback();
        $times[] = (hrtime(true) - $start) / 1e9;
        $peak = max($peak, memory_get_peak_usage(true));
    }

    return [
        'avg_s'   => round(array_sum($times) / count($times), 3),
        'peak_mb' => round($peak / 1048576, 1),
    ];
}
