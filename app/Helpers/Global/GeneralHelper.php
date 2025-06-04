<?php

use Carbon\Carbon;

if (! function_exists('appName')) {
    /**
     * Helper to grab the application name.
     *
     * @return mixed
     */
    function appName()
    {
        return config('app.name', 'Laravel Boilerplate');
    }
}

if (! function_exists('carbon')) {
    /**
     * Create a new Carbon instance from a time.
     *
     * @param $time
     * @return Carbon
     *
     * @throws Exception
     */
    function carbon($time)
    {
        return new Carbon($time);
    }
}

if (! function_exists('homeRoute')) {
    /**
     * Return the route to the "home" page depending on authentication/authorization status.
     *
     * @return string
     */
    function homeRoute()
    {
        if (auth()->check()) {
            if (auth()->user()->isAdmin()) {
                return 'frontend.user.account';
            }

            if (auth()->user()->isUser()) {
                return 'frontend.user.account';
            }
        }

        return 'frontend.index';
    }
}

if (! function_exists('displayDate')) {
    /**
     * Muestra la fecha en el idioma de la app usando Carbon (sin strftime).
     *
     * @param $date
     * @param string|null $format
     * @return string
     */
    function displayDate($date, $format = null)
    {
        if (!$date) return '';
        $carbon = $date instanceof \Carbon\Carbon ? $date : new \Carbon\Carbon($date);
        $locale = app()->getLocale() ?? 'es';
        $carbon->locale($locale);
        // Formato tipo: lunes, 26 de mayo de 2025, 11:31 AM
        $format = $format ?: 'l, d \d\e F \d\e Y, h:i A';
        return ucfirst($carbon->translatedFormat($format));
    }
}
