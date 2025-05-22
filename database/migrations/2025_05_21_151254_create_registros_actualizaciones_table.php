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

        Schema::create('registros_actualizaciones', function (Blueprint $table) {
            $table->integer('id');
            $table->enum('tipo_tabla', [""]);
            $table->string('nombre_archivo', 255);
            $table->dateTime('fecha_actualizacion')->useCurrent();
            $table->integer('registros_actualizados');
            $table->integer('registros_nuevos');
            $table->string('usuario_id', 79);
            $table->foreign('usuario_id')->references('numero_documento')->on('usuario');
        });

        Schema::enableForeignKeyConstraints();
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('registros_actualizaciones');
    }
};
