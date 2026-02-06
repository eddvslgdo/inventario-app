const ServiceSalidas = {

  procesarSalida: function(data) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);

      if (data.volumen_L <= 0) throw new Error("El volumen debe ser mayor a 0");

      // 1. Obtener lotes disponibles
      let lotes = getInventarioDisponible(
        data.producto_id, 
        data.presentacion_id, 
        data.ubicacion_id
      );

      // 2. Verificar stock total
      const stockTotal = lotes.reduce((sum, item) => sum + item.volumen, 0);
      if (stockTotal < data.volumen_L) {
        throw new Error(`Stock insuficiente. Tienes ${stockTotal}, pides ${data.volumen_L}`);
      }

      // 3. Ordenar FIFO (Primero caducidad, luego ingreso)
      lotes.sort((a, b) => {
        if (a.caducidad && b.caducidad) return new Date(a.caducidad) - new Date(b.caducidad);
        return new Date(a.fecha_ingreso) - new Date(b.fecha_ingreso);
      });

      let pendiente = data.volumen_L;
      let filasParaBorrar = [];

      // 4. Descontar
      for (let lote of lotes) {
        if (pendiente <= 0) break;

        let aDescontar = 0;
        if (lote.volumen <= pendiente) {
          aDescontar = lote.volumen;
          pendiente -= lote.volumen;
          filasParaBorrar.push(lote.fila);
        } else {
          aDescontar = pendiente;
          let saldo = lote.volumen - pendiente;
          pendiente = 0;
          actualizarVolumenFila(lote.fila, saldo);
        }

        guardarLogSalida({
          producto_id: data.producto_id,
          presentacion_id: data.presentacion_id,
          ubicacion_origen: data.ubicacion_id,
          volumen_L: aDescontar,
          lote: lote.lote,
          destino: 'Salida Web',
          comentario: 'FIFO'
        });
      }

      // 5. Limpieza de filas vacías (de abajo hacia arriba)
      filasParaBorrar.sort((a, b) => b - a);
      filasParaBorrar.forEach(fila => borrarFilaInventario(fila));

      return { success: true, mensaje: "Salida registrada (FIFO)" };

    } catch (e) {
      throw new Error(e.message);
    } finally {
      lock.releaseLock();
    }
  }
};