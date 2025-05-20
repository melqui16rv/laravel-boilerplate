@extends('backend.layouts.app')

@section('title', __('presupuesto'))

@section('content')
    <x-backend.card>
        <x-slot name="header">
            @lang('Welcome :Name', ['name' => $logged_in_user->name])
        </x-slot>

        <x-slot name="body">
            @lang('Bienvenido al panel presupuestal')
        </x-slot>
    </x-backend.card>
@endsection
