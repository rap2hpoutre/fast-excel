<?php

/*
 * FastExcel micro-benchmark.
 *
 * Usage:
 *   php bench/bench.php [rows] [runs] [--json]
 *   php bench/bench.php --compare base.json head.json
 *
 * Reports the median wall-clock time and peak memory for the main
 * export/import paths. Memory numbers are stable across runs; time on
 * shared CI runners is noisy, so deltas under ~10% should be ignored.
 *
 * Peak memory is read with memory_get_peak_usage(false) — PHP's own
 * allocator, which is released back between cases. The real (true) figure
 * reports the process high-water mark from the OS allocator, which never
 * shrinks, so a 4 MB streaming case measured after a 30 MB collection case
 * would wrongly report ~30 MB. Using the PHP figure keeps every case honest
 * in a single process.
 */

use OpenSpout\Common\Entity\Style\Style;
use Rap2hpoutre\FastExcel\FastExcel;

// Compare mode is pure PHP; only benchmark mode needs the autoloader.
if (($argv[1] ?? '') === '--compare') {
    exit(compareResults($argv[2], $argv[3]));
}

require getenv('BENCH_AUTOLOAD') ?: __DIR__.'/../vendor/autoload.php';

$rows = max(1, (int) ($argv[1] ?? 30000));
$runs = max(1, (int) ($argv[2] ?? 3));
$json = in_array('--json', $argv, true);

$dir = sys_get_temp_dir().'/fastexcel-bench-'.getmypid();
mkdir($dir);

$style = (new Style())->setFontBold();

$results = [];

// The collection cases share one materialized collection (15 MB+ at 30k rows).
// It is freed before the streaming cases so their peak memory reflects the
// streaming path alone, not a leftover collection still held in scope.
$collection = collect(array_map(fn ($i) => [
    'id'     => $i,
    'name'   => 'name_'.$i,
    'email'  => 'user'.$i.'@example.com',
    'amount' => $i * 1.5,
], range(1, $rows)));

$results['export_plain'] = bench($runs, function () use ($collection, $dir) {
    (new FastExcel($collection))->export($dir.'/plain.xlsx');
});
$results['export_styled'] = bench($runs, function () use ($collection, $dir, $style) {
    (new FastExcel($collection))->headerStyle($style)->rowsStyle($style)->export($dir.'/styled.xlsx');
});

unset($collection);
gc_collect_cycles();

// Streaming export: a generator never materializes the full dataset, so peak
// memory stays flat no matter how many rows are written.
$results['export_generator'] = bench($runs, function () use ($rows, $dir) {
    (new FastExcel(rowGenerator($rows)))->export($dir.'/gen.xlsx');
});

// Import cases read the file written by export_plain above.
$results['import'] = bench($runs, function () use ($dir) {
    (new FastExcel())->import($dir.'/plain.xlsx');
});
// Streaming import: importLazy() yields rows one at a time instead of
// accumulating them into a Collection.
$results['import_lazy'] = bench($runs, function () use ($dir) {
    (new FastExcel())->importLazy($dir.'/plain.xlsx')->each(fn ($row) => $row);
});
$results['import_transposed'] = bench($runs, function () use ($dir) {
    (new FastExcel())->transpose()->import($dir.'/plain.xlsx');
});

array_map('unlink', glob($dir.'/*'));
rmdir($dir);

if ($json) {
    echo json_encode(['rows' => $rows, 'runs' => $runs, 'php' => PHP_VERSION, 'results' => $results], JSON_PRETTY_PRINT).PHP_EOL;
} else {
    printf('%d rows, median of %d runs (PHP %s)%s', $rows, $runs, PHP_VERSION, PHP_EOL);
    foreach ($results as $name => $result) {
        printf('  %-18s %8.3fs  %8.1f MB peak%s', $name, $result['median_s'], $result['peak_mb'], PHP_EOL);
    }
}

/**
 * A fresh generator over the benchmark row shape. Generators are single-use,
 * so each timed run needs a new one.
 */
function rowGenerator(int $rows): Generator
{
    for ($i = 1; $i <= $rows; $i++) {
        yield [
            'id'     => $i,
            'name'   => 'name_'.$i,
            'email'  => 'user'.$i.'@example.com',
            'amount' => $i * 1.5,
        ];
    }
}

function bench(int $runs, callable $callback): array
{
    $times = [];
    $peak = 0;

    for ($i = 0; $i < $runs; $i++) {
        gc_collect_cycles();
        if (function_exists('memory_reset_peak_usage')) {
            memory_reset_peak_usage();
        }
        $start = hrtime(true);
        $callback();
        $times[] = (hrtime(true) - $start) / 1e9;
        $peak = max($peak, memory_get_peak_usage(false));
    }

    sort($times);

    return [
        'median_s' => round($times[intdiv(count($times) - 1, 2)], 3),
        'peak_mb'  => round($peak / 1048576, 1),
    ];
}

function compareResults(string $baseFile, string $headFile): int
{
    $base = json_decode(file_get_contents($baseFile), true);
    $head = json_decode(file_get_contents($headFile), true);

    echo "### Benchmark: base vs PR ({$head['rows']} rows, median of {$head['runs']} runs, PHP {$head['php']})\n\n";
    echo "| Benchmark | Base time | PR time | Δ time | Base peak mem | PR peak mem | Δ mem |\n";
    echo "|---|---|---|---|---|---|---|\n";

    $regression = false;

    foreach ($head['results'] as $name => $headResult) {
        $baseResult = $base['results'][$name] ?? null;
        if (!$baseResult) {
            printf("| %s | – | %.3fs | new | – | %.1f MB | new |\n", $name, $headResult['median_s'], $headResult['peak_mb']);
            continue;
        }
        $timeDelta = pct($baseResult['median_s'], $headResult['median_s']);
        $memDelta = pct($baseResult['peak_mb'], $headResult['peak_mb']);
        // Time on shared runners is noisy; only flag clear regressions.
        if ($timeDelta > 20 || $memDelta > 10) {
            $regression = true;
        }
        printf(
            "| %s | %.3fs | %.3fs | %+.1f%% | %.1f MB | %.1f MB | %+.1f%% |\n",
            $name,
            $baseResult['median_s'],
            $headResult['median_s'],
            $timeDelta,
            $baseResult['peak_mb'],
            $headResult['peak_mb'],
            $memDelta
        );
    }

    echo "\n_Time deltas under ~10% are runner noise; memory deltas are reliable._\n";

    if ($regression) {
        echo "\n⚠️ **Possible performance regression detected** (time +20% or memory +10%).\n";
    }

    return 0;
}

function pct(float $base, float $head): float
{
    return $base > 0 ? (($head - $base) / $base) * 100 : 0.0;
}
