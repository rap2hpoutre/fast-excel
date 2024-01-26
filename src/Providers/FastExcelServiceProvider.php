<?php

namespace Rap2hpoutre\FastExcel\Providers;

use Illuminate\Support\ServiceProvider;

class FastExcelServiceProvider extends ServiceProvider
{
    public const CONFIG_FILENAME = 'fast-excel';

    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot()
    {
        $this->publishes([
            $this->getConfigFile() => config_path(self::CONFIG_FILENAME . '.php'),
        ], 'config');
    }

    /**
     * Register any application services.
     *
     * @SuppressWarnings("unused")
     *
     * @return void
     */
    public function register()
    {
        $this->mergeConfigFrom(
            $this->getConfigFile(),
            self::CONFIG_FILENAME
        );

        $this->app->bind('fastexcel', function ($app, $data = null) {
            if (is_array($data)) {
                $data = collect($data);
            }

            return new \Rap2hpoutre\FastExcel\FastExcel($data);
        });
    }

    /**
     * @return string
     */
    protected function getConfigFile(): string
    {
        return __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'config' . DIRECTORY_SEPARATOR . self::CONFIG_FILENAME . '.php';
    }
}
