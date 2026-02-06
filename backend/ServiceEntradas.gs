const ServiceEntradas = {
  
  procesarEntrada: function(data) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000); 

      if (!data.producto_id || !data.ubicacion_id) throw new Error("Faltan datos clave.");
      if (data.volumen_L <= 0) throw new Error("Volumen incorrecto.");
      if (!data.lote) throw new Error("El Lote es obligatorio.");

      // --- CÁLCULO DE CADUCIDAD (+2 AÑOS) ---
      let fechaCaducidadCalculada = '';
      
      if (data.fecha_elaboracion) {
        // Truco seguro para sumar años sin problemas de zona horaria:
        // El input date viene como "2024-02-04"
        const partes = data.fecha_elaboracion.split('-'); // ["2024", "02", "04"]
        const anio = parseInt(partes[0]);
        const resto = partes.slice(1).join('-'); // "02-04"
        
        // Sumamos 2 al año y reconstruimos
        fechaCaducidadCalculada = (anio + 2) + '-' + resto;
      }
      // --------------------------------------

      const registro = {
        producto_id: data.producto_id,
        presentacion_id: data.presentacion_id,
        ubicacion_destino: data.ubicacion_id,
        ubicacion_id: data.ubicacion_id,
        volumen_L: Number(data.volumen_L),
        lote: data.lote.toUpperCase(),
        
        // Guardamos las fechas procesadas
        fecha_elaboracion: data.fecha_elaboracion || '',
        fecha_caducidad: fechaCaducidadCalculada, // <--- Aquí va el dato calculado
        
        proveedor: data.proveedor || '', // <--- Guardamos proveedor
        comentario: 'Entrada Web'
      };

      guardarLogEntrada(registro);
      actualizarInventarioEntrada(registro);

      return { success: true, mensaje: "Entrada registrada" };

    } catch (e) {
      throw new Error(e.message);
    } finally {
      lock.releaseLock();
    }
  }
};