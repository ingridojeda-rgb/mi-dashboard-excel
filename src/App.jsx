import { useState } from "react";
import * as XLSX from "xlsx";

function App() {
  const [tablaTipificaciones, setTablaTipificaciones] = useState([]);
  const [tablaCalificacion, setTablaCalificacion] = useState([]);
  const [gruposEstado, setGruposEstado] = useState([]);

  const [totalTipificaciones, setTotalTipificaciones] = useState(0);
  const [totalCalificacion, setTotalCalificacion] = useState(0);
  const [totalGeneralEstado, setTotalGeneralEstado] = useState(0);

  const [cargando, setCargando] = useState(false);
  const [nombreArchivo, setNombreArchivo] = useState("");
  const [error, setError] = useState("");

  const manejarSubidaArchivo = (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    setNombreArchivo(archivo.name);
    setCargando(true);
    setError("");

    const lector = new FileReader();

    lector.onload = (evento) => {
      try {
        const buffer = evento.target.result;
        const libro = XLSX.read(buffer, { type: "array" });

        const nombreHoja = libro.SheetNames[0];
        const hoja = libro.Sheets[nombreHoja];

        const datosJson = XLSX.utils.sheet_to_json(hoja, { defval: "" });

        if (datosJson.length === 0) {
          setError("El archivo Excel no contiene filas de datos.");
          setCargando(false);
          return;
        }

        procesarMetricas(datosJson);
      } catch (err) {
        console.error(err);
        setError("Error al leer el archivo Excel. Asegúrate de que sea un archivo .xlsx o .xls válido.");
      } finally {
        setCargando(false);
      }
    };

    lector.readAsArrayBuffer(archivo);
  };

  const procesarMetricas = (datos) => {
    const conteoTipificaciones = {};
    const conteoCalificacion = {};
    const mapaEstados = {};

    let sumaTotalTip = 0;
    let sumaTotalCal = 0;
    let sumaTotalEst = 0;

    datos.forEach((fila) => {
      const valTip = fila["SD - Tipificaciones"];
      const tipificacion = (valTip !== undefined && valTip !== null && valTip.toString().trim() !== "")
        ? valTip.toString().trim()
        : "Sin especificar / Vacío";

      conteoTipificaciones[tipificacion] = (conteoTipificaciones[tipificacion] || 0) + 1;
      sumaTotalTip += 1;

      const valCal = fila["SD - Calidad - Calificación"];
      const calificacion = (valCal !== undefined && valCal !== null && valCal.toString().trim() !== "")
        ? valCal.toString().trim()
        : "Sin especificar / Vacío";

      conteoCalificacion[calificacion] = (conteoCalificacion[calificacion] || 0) + 1;
      sumaTotalCal += 1;

      const valEstado = fila["Estado"];
      const estado = (valEstado !== undefined && valEstado !== null && valEstado.toString().trim() !== "")
        ? valEstado.toString().trim()
        : "Sin especificar / Vacío";

      if (!mapaEstados[estado]) {
        mapaEstados[estado] = {
          nombreEstado: estado,
          totalEstado: 0,
          motivosMap: {}
        };
      }

      mapaEstados[estado].totalEstado += 1;
      mapaEstados[estado].motivosMap[calificacion] = (mapaEstados[estado].motivosMap[calificacion] || 0) + 1;
      sumaTotalEst += 1;
    });

    const formatoTablaTip = Object.entries(conteoTipificaciones)
      .map(([categoria, cantidad]) => ({
        categoria,
        cantidad,
        porcentaje: sumaTotalTip > 0 ? ((cantidad / sumaTotalTip) * 100).toFixed(2) + "%" : "0%"
      }))
      .sort((a, b) => b.cantidad - a.cantidad);

    const formatoTablaCal = Object.entries(conteoCalificacion)
      .map(([categoria, cantidad]) => ({
        categoria,
        cantidad,
        porcentaje: sumaTotalCal > 0 ? ((cantidad / sumaTotalCal) * 100).toFixed(2) + "%" : "0%"
      }))
      .sort((a, b) => b.cantidad - a.cantidad);

    const estructuraGrupos = Object.values(mapaEstados)
      .map((itemEstado) => {
        const listaMotivos = Object.entries(itemEstado.motivosMap)
          .map(([motivo, cant]) => ({
            motivo,
            cantidad: cant,
            porcentaje: sumaTotalEst > 0 ? ((cant / sumaTotalEst) * 100).toFixed(2) + "%" : "0%"
          }))
          .sort((a, b) => b.cantidad - a.cantidad);

        return {
          estado: itemEstado.nombreEstado,
          totalEstado: itemEstado.totalEstado,
          porcentajeEstado: sumaTotalEst > 0 ? ((itemEstado.totalEstado / sumaTotalEst) * 100).toFixed(2) + "%" : "0%",
          motivos: listaMotivos
        };
      })
      .sort((a, b) => b.totalEstado - a.totalEstado);

    setTablaTipificaciones(formatoTablaTip);
    setTablaCalificacion(formatoTablaCal);
    setGruposEstado(estructuraGrupos);

    setTotalTipificaciones(sumaTotalTip);
    setTotalCalificacion(sumaTotalCal);
    setTotalGeneralEstado(sumaTotalEst);
  };

  const estilos = {
    contenedor: {
      fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
      backgroundColor: "#f1f5f9",
      minHeight: "100vh",
      padding: "40px 20px",
      color: "#334155"
    },
    header: {
      textAlign: "center",
      marginBottom: "30px"
    },
    titulo: {
      color: "#1e3a8a",
      fontSize: "30px",
      fontWeight: "bold",
      marginBottom: "8px"
    },
    subtitulo: {
      color: "#64748b",
      fontSize: "15px"
    },
    zonaCarga: {
      background: "#ffffff",
      border: "2px dashed #94a3b8",
      borderRadius: "12px",
      padding: "30px",
      textAlign: "center",
      maxWidth: "700px",
      margin: "0 auto 40px auto",
      boxShadow: "0 4px 6px -1px rgba(0, 0, 0, 0.05)"
    },
    inputArchivo: {
      marginTop: "15px",
      padding: "8px"
    },
    seccionVertical: {
      display: "flex",
      flexDirection: "column",
      gap: "35px",
      maxWidth: "1400px",
      margin: "0 auto"
    },
    tarjeta: {
      background: "#ffffff",
      borderRadius: "10px",
      padding: "25px",
      boxShadow: "0 4px 6px -1px rgba(0, 0, 0, 0.1)",
      width: "100%",
      boxSizing: "border-box"
    },
    tituloTabla: {
      color: "#1e3a8a",
      fontSize: "20px",
      borderBottom: "3px solid #1e3a8a",
      paddingBottom: "8px",
      marginTop: 0,
      marginBottom: "20px"
    },
    tabla: {
      width: "100%",
      borderCollapse: "collapse"
    },
    tablaPivot: {
      width: "100%",
      borderCollapse: "collapse",
      border: "1px solid #cbd5e1"
    },
    thPivot: {
      backgroundColor: "#475569",
      color: "#ffffff",
      padding: "10px 14px",
      textAlign: "left",
      fontSize: "14px",
      border: "1px solid #cbd5e1"
    },
    tdPivot: {
      padding: "8px 14px",
      border: "1px solid #e2e8f0",
      fontSize: "13px",
      color: "#1e293b",
      whiteSpace: "normal",
      wordBreak: "break-word"
    },
    rowSubtotal: {
      backgroundColor: "#e2e8f0",
      fontWeight: "bold",
      color: "#0f172a"
    },
    rowTotalGeneral: {
      backgroundColor: "#cbd5e1",
      fontWeight: "bold",
      fontSize: "14px",
      color: "#0f172a"
    },
    th: {
      backgroundColor: "#1e3a8a",
      color: "#ffffff",
      padding: "12px 16px",
      textAlign: "left",
      fontSize: "14px"
    },
    td: {
      padding: "10px 16px",
      borderBottom: "1px solid #e2e8f0",
      fontSize: "13px",
      whiteSpace: "normal",
      wordBreak: "break-word"
    },
    tfTd: {
      padding: "12px 16px",
      borderTop: "2px solid #1e3a8a",
      fontSize: "14px",
      fontWeight: "bold",
      backgroundColor: "#f1f5f9",
      color: "#1e3a8a"
    },
    iconoContraer: {
      display: "inline-block",
      backgroundColor: "#94a3b8",
      color: "white",
      fontWeight: "bold",
      borderRadius: "2px",
      width: "14px",
      height: "14px",
      lineHeight: "12px",
      textAlign: "center",
      fontSize: "12px",
      marginRight: "8px"
    },
    mensajeVacio: {
      textAlign: "center",
      color: "#94a3b8",
      marginTop: "50px",
      fontSize: "16px"
    }
  };

  return (
    <div style={estilos.contenedor}>
      <header style={estilos.header}>
        <h1 style={estilos.titulo}>📊 Escalados a funcional</h1>
        <p style={estilos.subtitulo}>Carga tu base de datos en Excel para generar las métricas automáticamente</p>
      </header>

      <div style={estilos.zonaCarga}>
        <label style={{ fontWeight: "bold", display: "block" }}>
          📁 Selecciona o arrastra tu archivo Excel:
        </label>
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={manejarSubidaArchivo}
          style={estilos.inputArchivo}
        />
        {nombreArchivo && (
          <p style={{ color: "#2563eb", fontWeight: "bold", marginTop: "10px" }}>
            Archivo cargado: {nombreArchivo}
          </p>
        )}
        {error && <p style={{ color: "#dc2626", marginTop: "10px" }}>{error}</p>}
      </div>

      {cargando && (
        <p style={{ textAlign: "center", fontSize: "18px", color: "#1e3a8a" }}>
          Procesando datos del Excel...
        </p>
      )}

      {tablaTipificaciones.length > 0 && (
        <div style={estilos.seccionVertical}>
          {/* TABLA 1: SD - Tipificaciones */}
          <div style={estilos.tarjeta}>
            <h2 style={estilos.tituloTabla}>1. SD - Tipificaciones</h2>
            <table style={estilos.tabla}>
              <thead>
                <tr>
                  <th style={estilos.th}>Tipificación</th>
                  <th style={{ ...estilos.th, textAlign: "right", width: "140px" }}>Cantidad</th>
                  <th style={{ ...estilos.th, textAlign: "right", width: "120px" }}>%</th>
                </tr>
              </thead>
              <tbody>
                {tablaTipificaciones.map((item, index) => (
                  <tr key={index} style={{ backgroundColor: index % 2 === 0 ? "#ffffff" : "#f8fafc" }}>
                    <td style={estilos.td}>{item.categoria}</td>
                    <td style={{ ...estilos.td, textAlign: "right", fontWeight: "bold" }}>
                      {item.cantidad}
                    </td>
                    <td style={{ ...estilos.td, textAlign: "right", color: "#2563eb", fontWeight: "bold" }}>
                      {item.porcentaje}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td style={estilos.tfTd}>TOTAL</td>
                  <td style={{ ...estilos.tfTd, textAlign: "right" }}>{totalTipificaciones}</td>
                  <td style={{ ...estilos.tfTd, textAlign: "right" }}>100.00%</td>
                </tr>
              </tfoot>
            </table>
          </div>

          {/* TABLA 2: Motivos de escalado */}
          <div style={estilos.tarjeta}>
            <h2 style={estilos.tituloTabla}>2. Motivos de escalado</h2>
            <table style={estilos.tabla}>
              <thead>
                <tr>
                  <th style={estilos.th}>Motivo</th>
                  <th style={{ ...estilos.th, textAlign: "right", width: "140px" }}>Cantidad</th>
                  <th style={{ ...estilos.th, textAlign: "right", width: "120px" }}>%</th>
                </tr>
              </thead>
              <tbody>
                {tablaCalificacion.map((item, index) => (
                  <tr key={index} style={{ backgroundColor: index % 2 === 0 ? "#ffffff" : "#f8fafc" }}>
                    <td style={estilos.td}>{item.categoria}</td>
                    <td style={{ ...estilos.td, textAlign: "right", fontWeight: "bold" }}>
                      {item.cantidad}
                    </td>
                    <td style={{ ...estilos.td, textAlign: "right", color: "#2563eb", fontWeight: "bold" }}>
                      {item.porcentaje}
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td style={estilos.tfTd}>TOTAL</td>
                  <td style={{ ...estilos.tfTd, textAlign: "right" }}>{totalCalificacion}</td>
                  <td style={{ ...estilos.tfTd, textAlign: "right" }}>100.00%</td>
                </tr>
              </tfoot>
            </table>
          </div>

          {/* TABLA 3: Resumen por Estado */}
          <div style={estilos.tarjeta}>
            <h2 style={estilos.tituloTabla}>3. Resumen por Estado</h2>
            <table style={estilos.tablaPivot}>
              <thead>
                <tr>
                  <th style={{ ...estilos.thPivot, width: "250px" }}>Estado</th>
                  <th style={estilos.thPivot}>Motivo de escalado</th>
                  <th style={{ ...estilos.thPivot, textAlign: "right", width: "140px" }}>Cantidad</th>
                  <th style={{ ...estilos.thPivot, textAlign: "right", width: "120px" }}>%</th>
                </tr>
              </thead>
              <tbody>
                {gruposEstado.map((grupo, gIndex) => (
                  <tr key={`group-${gIndex}`}>
                    <td colSpan={4} style={{ padding: 0, border: 0 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <tbody>
                          {grupo.motivos.map((mItem, mIndex) => (
                            <tr key={`${gIndex}-${mIndex}`}>
                              <td style={{ ...estilos.tdPivot, width: "250px" }}>
                                {mIndex === 0 && (
                                  <span style={{ fontWeight: "600" }}>
                                    <span style={estilos.iconoContraer}>-</span>
                                    {grupo.estado}
                                  </span>
                                )}
                              </td>
                              <td style={estilos.tdPivot}>{mItem.motivo}</td>
                              <td style={{ ...estilos.tdPivot, textAlign: "right", width: "140px" }}>{mItem.cantidad}</td>
                              <td style={{ ...estilos.tdPivot, textAlign: "right", width: "120px" }}>{mItem.porcentaje}</td>
                            </tr>
                          ))}
                          <tr style={estilos.rowSubtotal}>
                            <td style={estilos.tdPivot} colSpan={2}>
                              Total {grupo.estado}
                            </td>
                            <td style={{ ...estilos.tdPivot, textAlign: "right", width: "140px" }}>{grupo.totalEstado}</td>
                            <td style={{ ...estilos.tdPivot, textAlign: "right", width: "120px" }}>{grupo.porcentajeEstado}</td>
                          </tr>
                        </tbody>
                      </table>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={estilos.rowTotalGeneral}>
                  <td style={{ ...estilos.tdPivot, borderTop: "2px solid #0f172a" }} colSpan={2}>
                    Suma total
                  </td>
                  <td style={{ ...estilos.tdPivot, textAlign: "right", borderTop: "2px solid #0f172a", width: "140px" }}>
                    {totalGeneralEstado}
                  </td>
                  <td style={{ ...estilos.tdPivot, textAlign: "right", borderTop: "2px solid #0f172a", width: "120px" }}>
                    100.00%
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {tablaTipificaciones.length === 0 && !cargando && (
        <div style={estilos.mensajeVacio}>
          Aún no se han cargado datos. Por favor, selecciona un archivo Excel arriba.
        </div>
      )}
    </div>
  );
}

export default App;