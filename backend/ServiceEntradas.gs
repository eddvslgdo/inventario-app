const ServiceEntradas = {
  
  procesarEntrada: function(data) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000); 

      // --- VALIDACIONES ---
      if (!data.producto_id || !data.ubicacion_id) throw new Error("Faltan datos clave.");
      if (!(data.volumen_L > 0)) throw new Error("Volumen incorrecto.");
      if (!data.lote) throw new Error("El Lote es obligatorio.");

      // --- CÁLCULO DE FECHAS ---
      let elabTxt = "";
      let cadTxt = "";

      if (data.fecha_elaboracion && data.fecha_elaboracion.includes("-")) {
          const partes = data.fecha_elaboracion.split("-"); 
          const anio = parseInt(partes[0]); 
          const resto = partes[1] + "-" + partes[2]; 
          elabTxt = data.fecha_elaboracion;
          cadTxt = (anio + 2) + "-" + resto;
      } else {
          const hoy = new Date();
          const y = hoy.getFullYear();
          const m = ("0" + (hoy.getMonth() + 1)).slice(-2);
          const d = ("0" + hoy.getDate()).slice(-2);
          elabTxt = `${y}-${m}-${d}`;
          cadTxt = `${y + 2}-${m}-${d}`;
      }

      // --- OBJETO REGISTRO (CAMBIO CLAVE 1) ---
      // AQUI AGREGAMOS LAS FECHAS AL OBJETO PARA QUE VIAJEN JUNTAS
      const registro = {
        producto_id: data.producto_id,
        presentacion_id: data.presentacion_id,
        ubicacion_destino: data.ubicacion_id,
        ubicacion_id: data.ubicacion_id,
        volumen_L: Number(data.volumen_L),
        lote: data.lote.toUpperCase(),
        proveedor: data.proveedor || '',
        comentario: 'Entrada Web',
        fecha_caducidad: cadTxt,    // <--- ¡NUEVO!
        fecha_elaboracion: elabTxt  // <--- ¡NUEVO!
      };

      // 1. Guardamos el Log
      guardarLogEntrada(registro);

      // 2. Actualizamos Inventario
      // (CAMBIO CLAVE 2: ELIMINAMOS TODO EL BLOQUE DEL "ESTAMPADO QUIRÚRGICO")
      actualizarInventarioEntrada(registro);

      return { success: true, mensaje: "Entrada registrada correctamente" };

    } catch (e) {
      console.error(e);
      throw new Error(e.message);
    } finally {
      lock.releaseLock();
    }
  }
};