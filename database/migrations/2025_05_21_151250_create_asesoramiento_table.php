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
        Schema::create('asesoramiento', function (Blueprint $table) {
            $table->integer('id_asesoramiendo')->primary();
            $table->enum('tipo', [""]);
            $table->string('encargadoAsesoramiento', 155);
            $table->string('nombreEntidadImpacto', 155);
            $table->dateTime('fechaAsesoramiento');
            $table->dateTime('fecha_insert')->useCurrent();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('asesoramiento');
    }
};
