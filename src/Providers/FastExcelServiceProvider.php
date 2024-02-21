<?php

namespace Rap2hpoutre\FastExcel\Providers;

use Illuminate\Support\ServiceProvider;
use Illuminate\Contracts\Support\DeferrableProvider;

class FastExcelServiceProvider extends ServiceProvider implements DeferrableProvider
{
    /**
     * Register any application services.
     *
     * @SuppressWarnings("unused")
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind('fastexcel', function ($app, $data = null) {
            if (is_array($data)) {
                $data = collect($data);
            }

            return new \Rap2hpoutre\FastExcel\FastExcel($data);
        });
    }

    /**
     * Get the services provided by the provider.
     *
     * @return array
     */
    public function provides(): array
    {
        return [
            'fastexcel',
        ];
    }
}
