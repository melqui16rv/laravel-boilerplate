<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::disableForeignKeyConstraints();

        Schema::create('imagenes_saldos_asignados', function (Blueprint $table) {
            $table->integer('ID_IMAGEN');
            $table->integer('ID_SALDO');
            $table->foreign('ID_SALDO')->references('ID_SALDO')->on('saldos_asignados');
            $table->string('NOMBRE_ORIGINAL', 255);
            $table->string('RUTA_IMAGEN', 255);
            $table->dateTime('FECHA_SUBIDA')->useCurrent();
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('imagenes_saldos_asignados');
    }
};
