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

        Schema::create('saldos_asignados', function (Blueprint $table) {
            $table->integer('ID_SALDO');
            $table->string('NOMBRE_PERSONA', 255);
            $table->string('DOCUMENTO_PERSONA', 55);
            $table->dateTime('FECHA_REGISTRO')->useCurrent();
            $table->date('FECHA_INICIO');
            $table->date('FECHA_FIN');
            $table->date('FECHA_PAGO')->nullable()->default('DEFAULT NULL');
            $table->decimal('SALDO_ASIGNADO', 15, 2);
            $table->string('CODIGO_CRP', 55);
            $table->string('CODIGO_CDP', 55);
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('saldos_asignados');
    }
};
