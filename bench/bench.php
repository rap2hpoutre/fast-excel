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

$collection = collect(array_map(fn ($i) => [
    'id'     => $i,
    'name'   => 'name_'.$i,
    'email'  => 'user'.$i.'@example.com',
    'amount' => $i * 1.5,
], range(1, $rows)));

$style = (new Style())->setFontBold();

$results = [
    'export_plain' => bench($runs, function () use ($collection, $dir) {
        (new FastExcel($collection))->export($dir.'/plain.xlsx');
    }),
    'export_styled' => bench($runs, function () use ($collection, $dir, $style) {
        (new FastExcel($collection))->headerStyle($style)->rowsStyle($style)->export($dir.'/styled.xlsx');
    }),
    'import' => bench($runs, function () use ($dir) {
        (new FastExcel())->import($dir.'/plain.xlsx');
    }),
    'import_transposed' => bench($runs, function () use ($dir) {
        (new FastExcel())->transpose()->import($dir.'/plain.xlsx');
    }),
];

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
        $peak = max($peak, memory_get_peak_usage(true));
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
