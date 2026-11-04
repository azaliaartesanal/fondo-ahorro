'use client';

import { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import * as XLSX from 'xlsx';
import { Check, X, Download, Upload, ShieldCheck, Users, RefreshCw } from 'lucide-react';

interface Transaccion {
  id: number;
  monto_reportado: number;
  concepto: string;
  referencia_bre_b: string;
  comprobante_url: string;
  estado: string;
  fecha_reporte: string;
  socios: { nombre: string };
}

export default function AdminPage() {
  const [autorizado, setAutorizado] = useState(false);
  const [pin, setPin] = useState('');
  const [transacciones, setTransacciones] = useState<Transaccion[]>([]);
  const [cargando, setCargando] = useState(false);
  const [estadoSubida, setEstadoSubida] = useState<string | null>(null);

  // Define tu clave de acceso aquí
  const PIN_CORRECTO = '1234';

  useEffect(() => {
    if (autorizado) {
      cargarTransacciones();
    }
  }, [autorizado]);

  async function cargarTransacciones() {
    setCargando(true);
    const { data, error } = await supabase
      .from('transacciones')
      .select('id, monto_reportado, concepto, referencia_bre_b, comprobante_url, estado, fecha_reporte, socios(nombre)')
      .order('fecha_reporte', { ascending: false });

    if (!error && data) {
      setTransacciones(data as any);
    }
    setCargando(false);
  }

  // Carga masiva de socios desde un archivo Excel (.xlsx o .csv)
  async function procesarExcelSocios(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    setEstadoSubida('Leyendo archivo Excel...');
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const filas: any[] = XLSX.utils.sheet_to_json(ws);

        // Detecta nombres desde columnas como 'nombre', 'Nombre', 'socio' o 'Socio'
        const sociosParaInsertar = filas
          .map((f) => ({
            nombre: String(f.nombre || f.Nombre || f.socio || f.Socio || f.NOMBRE || '').trim(),
            telefono: String(f.telefono || f.Telefono || f.celular || f.Celular || '').trim(),
          }))
          .filter((s) => s.nombre.length > 0);

        if (sociosParaInsertar.length === 0) {
          setEstadoSubida('Error: No se encontró una columna llamada "nombre" en el Excel.');
          return;
        }

        setEstadoSubida(`Guardando ${sociosParaInsertar.length} socios en la base de datos...`);
        const { error } = await supabase.from('socios').upsert(sociosParaInsertar, { onConflict: 'nombre' });

        if (error) throw error;
        setEstadoSubida(`¡Éxito! Se cargaron ${sociosParaInsertar.length} socios correctamente.`);
      } catch (err: any) {
        setEstadoSubida('Error: ' + err.message);
      }
    };

    reader.readAsBinaryString(file);
  }

  async function actualizarEstado(id: number, nuevoEstado: 'aprobado' | 'rechazado') {
    const { error } = await supabase
      .from('transacciones')
      .update({ estado: nuevoEstado, fecha_aprobacion: new Date().toISOString() })
      .eq('id', id);

    if (!error) {
      setTransacciones((prev) =>
        prev.map((t) => (t.id === id ? { ...t, estado: nuevoEstado } : t))
      );
    }
  }

  // Exportar histórico a Excel
  function exportarAExcel() {
    const datosFormateados = transacciones.map((t) => ({
      ID: t.id,
      Fecha: new Date(t.fecha_reporte).toLocaleDateString(),
      Socio: t.socios?.nombre || 'Desconocido',
      Monto: t.monto_reportado,
      Concepto: t.concepto,
      Referencia_BreB: t.referencia_bre_b,
      Estado: t.estado,
      Comprobante: t.comprobante_url,
    }));

    const ws = XLSX.utils.json_to_sheet(datosFormateados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Auditoria_Fondo');
    XLSX.writeFile(wb, `Reporte_Fondo_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // Pantalla de bloqueo si no ha puesto el PIN
  if (!autorizado) {
    return (
      <main className="min-h-screen bg-slate-900 text-white flex items-center justify-center p-4">
        <div className="bg-slate-800 p-6 rounded-2xl border border-slate-700 max-w-sm w-full space-y-4 text-center">
          <div className="inline-block p-3 bg-emerald-500/20 text-emerald-400 rounded-full">
            <ShieldCheck className="w-8 h-8" />
          </div>
          <h1 className="text-xl font-bold">Auditoría del Fondo</h1>
          <p className="text-sm text-slate-400">Ingresa la clave para acceder</p>
          <input
            type="password"
            placeholder="Clave de acceso"
            value={pin}
            onChange={(e) => setPin(e.target.value)}
            className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-center text-white text-lg tracking-widest focus:outline-none focus:border-emerald-500"
          />
          <button
            onClick={() => {
              if (pin === PIN_CORRECTO) setAutorizado(true);
              else alert('Clave incorrecta');
            }}
            className="w-full bg-emerald-600 hover:bg-emerald-500 font-semibold py-2.5 rounded-lg transition duration-200"
          >
            Entrar
          </button>
        </div>
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-slate-900 text-slate-100 p-4 md:p-8">
      <div className="max-w-6xl mx-auto space-y-6">
        {/* Cabecera y acciones */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-800 p-6 rounded-2xl border border-slate-700">
          <div>
            <h1 className="text-2xl font-bold text-white flex items-center gap-2">
              <ShieldCheck className="w-7 h-7 text-emerald-400" /> Auditoría y Conciliación
            </h1>
            <p className="text-sm text-slate-400 mt-1">Cruce de pagos Bre-B y administración</p>
          </div>

          <div className="flex flex-wrap gap-3">
            {/* Carga de Excel de Socios */}
            <label className="bg-blue-600 hover:bg-blue-500 text-white px-4 py-2.5 rounded-xl cursor-pointer flex items-center gap-2 text-sm font-medium transition duration-200">
              <Users className="w-4 h-4" /> Importar Socios (.xlsx)
              <input type="file" accept=".xlsx, .xls, .csv" onChange={procesarExcelSocios} className="hidden" />
            </label>

            {/* Descarga a Excel */}
            <button
              onClick={exportarAExcel}
              className="bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2.5 rounded-xl flex items-center gap-2 text-sm font-medium transition duration-200"
            >
              <Download className="w-4 h-4" /> Exportar a Excel
            </button>

            {/* Refrescar */}
            <button
              onClick={cargarTransacciones}
              className="bg-slate-700 hover:bg-slate-600 text-slate-300 p-2.5 rounded-xl transition duration-200"
              title="Refrescar lista"
            >
              <RefreshCw className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Notificación de carga de socios */}
        {estadoSubida && (
          <div className="p-3 bg-blue-500/10 border border-blue-500/30 rounded-xl text-blue-300 text-sm">
            {estadoSubida}
          </div>
        )}

        {/* Tabla de pagos reportados */}
        <div className="bg-slate-800 rounded-2xl border border-slate-700 overflow-hidden shadow-xl">
          <div className="p-4 border-b border-slate-700 font-semibold text-slate-200">
            Comprobantes Pendientes y Reportados
          </div>

          {cargando ? (
            <div className="p-8 text-center text-slate-400">Cargando pagos...</div>
          ) : transacciones.length === 0 ? (
            <div className="p-8 text-center text-slate-400">Aún no hay comprobantes reportados.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm text-slate-300">
                <thead className="bg-slate-700/50 text-slate-400 uppercase text-xs">
                  <tr>
                    <th className="p-4">Fecha</th>
                    <th className="p-4">Socio</th>
                    <th className="p-4">Monto</th>
                    <th className="p-4">Concepto</th>
                    <th className="p-4">Ref Bre-B</th>
                    <th className="p-4">Comprobante</th>
                    <th className="p-4">Estado</th>
                    <th className="p-4 text-center">Acción</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/50">
                  {transacciones.map((t) => (
                    <tr key={t.id} className="hover:bg-slate-700/30">
                      <td className="p-4">{new Date(t.fecha_reporte).toLocaleDateString()}</td>
                      <td className="p-4 font-medium text-white">{t.socios?.nombre || 'Desconocido'}</td>
                      <td className="p-4 text-emerald-400 font-bold">${Number(t.monto_reportado).toLocaleString()}</td>
                      <td className="p-4 capitalize">{t.concepto}</td>
                      <td className="p-4 font-mono text-xs">{t.referencia_bre_b || 'N/A'}</td>
                      <td className="p-4">
                        {t.comprobante_url ? (
                          <a
                            href={t.comprobante_url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-blue-400 underline hover:text-blue-300 text-xs"
                          >
                            Ver Imagen
                          </a>
                        ) : (
                          'Sin archivo'
                        )}
                      </td>
                      <td className="p-4">
                        <span
                          className={`px-2.5 py-1 rounded-full text-xs font-semibold ${
                            t.estado === 'aprobado'
                              ? 'bg-emerald-500/20 text-emerald-400'
                              : t.estado === 'rechazado'
                              ? 'bg-red-500/20 text-red-400'
                              : 'bg-amber-500/20 text-amber-400'
                          }`}
                        >
                          {t.estado}
                        </span>
                      </td>
                      <td className="p-4">
                        <div className="flex justify-center gap-2">
                          <button
                            onClick={() => actualizarEstado(t.id, 'aprobado')}
                            className="p-1.5 bg-emerald-600/30 hover:bg-emerald-600 text-emerald-300 hover:text-white rounded-lg transition"
                            title="Aprobar pago"
                          >
                            <Check className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => actualizarEstado(t.id, 'rechazado')}
                            className="p-1.5 bg-red-600/30 hover:bg-red-600 text-red-300 hover:text-white rounded-lg transition"
                            title="Rechazar pago"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </main>
  );
}