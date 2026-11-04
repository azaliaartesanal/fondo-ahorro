'use client';

import { useState, useEffect } from 'react';
import { supabase } from './lib/supabase';
import * as XLSX from 'xlsx';
import { 
  AlertCircle, 
  CheckCircle2, 
  Upload, 
  Wallet, 
  ShieldCheck, 
  Users, 
  RefreshCw, 
  Check, 
  X, 
  Lock, 
  LogOut, 
  CreditCard, 
  PlusCircle, 
  FileSpreadsheet,
  Tag,
  CalendarCheck,
  Trash2,
  AlertTriangle,
  PieChart,
  Landmark,
  Coins,
  Download,
  FileText,
  UserCheck,
  Filter,
  Search,
  Receipt,
  MinusCircle
} from 'lucide-react';

interface Socio {
  id: number;
  nombre: string;
  telefono?: string;
}

interface Transaccion {
  id: number;
  socio_id?: number;
  monto_reportado: number;
  concepto: string;
  pago_cuota_1?: boolean;
  pago_cuota_2?: boolean;
  desglose_ahorro: number;
  desglose_prestamo_capital: number;
  desglose_prestamo_interes: number;
  desglose_multa: number;
  desglose_actividad: number;
  comprobante_url: string;
  estado: string;
  fecha_reporte: string;
  socios: { nombre: string };
}

interface Prestamo {
  id: number;
  socio_id: number;
  monto_prestado: number;
  saldo_capital: number;
  tasa_interes_mensual: number;
  fecha_desembolso: string;
  estado: string;
  socios?: { nombre: string };
}

interface Actividad {
  id: number;
  nombre: string;
  valor_cuota: number;
  costo_premios?: number;
  fecha_limite: string;
  activa: boolean;
}

interface MovimientoCaja {
  id: number;
  tipo: 'ingreso_efectivo' | 'gasto_fondo';
  concepto: string;
  monto: number;
  fecha: string;
}

const MESES = [
  { valor: '01', nombre: 'Enero' },
  { valor: '02', nombre: 'Febrero' },
  { valor: '03', nombre: 'Marzo' },
  { valor: '04', nombre: 'Abril' },
  { valor: '05', nombre: 'Mayo' },
  { valor: '06', nombre: 'Junio' },
  { valor: '07', nombre: 'Julio' },
  { valor: '08', nombre: 'Agosto' },
  { valor: '09', nombre: 'Septiembre' },
  { valor: '10', nombre: 'Octubre' },
  { valor: '11', nombre: 'Noviembre' },
  { valor: '12', nombre: 'Diciembre' }
];

function formatearMoneda(valor: number): string {
  if (isNaN(valor)) return '0';
  return Math.round(valor).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function esFechaExtemporanea(fechaStr: string): boolean {
  if (!fechaStr) return false;
  const partes = fechaStr.split('T')[0].split('-');
  const dia = parseInt(partes[2], 10);
  return (dia > 7 && dia < 15) || dia > 21;
}

export default function Home() {
  const [esAdmin, setEsAdmin] = useState(false);
  const [modoPublico, setModoPublico] = useState<'reportar' | 'consultar'>('reportar');
  const [pestanaAdmin, setPestanaAdmin] = useState<'balance' | 'comprobantes' | 'prestamos' | 'actividades' | 'atrasados'>('balance');
  const [mostrarModalPin, setMostrarModalPin] = useState(false);
  const [pinIngresado, setPinIngresado] = useState('');
  const PIN_CORRECTO = '1234';

  const [mostrarModalBorrado, setMostrarModalBorrado] = useState(false);
  const [pinBorrado, setPinBorrado] = useState('');
  const [cargandoBorrado, setCargandoBorrado] = useState(false);

  // Filtros de Comprobantes
  const [filtroComprobanteSocio, setFiltroComprobanteSocio] = useState<string>('todos');
  const [filtroComprobanteMes, setFiltroComprobanteMes] = useState<string>('todos');

  // Selector Mes/Año para Atrasados y Exportación
  const [mesSeleccionado, setMesSeleccionado] = useState('02');
  const [anioSeleccionado, setAnioSeleccionado] = useState('2027');
  const [socioParaExtracto, setSocioParaExtracto] = useState<number | null>(null);

  // Estados Globales
  const [socios, setSocios] = useState<Socio[]>([]);
  const [transacciones, setTransacciones] = useState<Transaccion[]>([]);
  const [prestamos, setPrestamos] = useState<Prestamo[]>([]);
  const [actividades, setActividades] = useState<Actividad[]>([]);
  const [movimientosCaja, setMovimientosCaja] = useState<MovimientoCaja[]>([]);

  // Formulario Socio
  const [socioSeleccionado, setSocioSeleccionado] = useState<number | null>(null);
  const [fechaPago, setFechaPago] = useState('');
  
  const [pagarAhorro, setPagarAhorro] = useState(true);
  const [cuota1Seleccionada, setCuota1Seleccionada] = useState(true);
  const [cuota2Seleccionada, setCuota2Seleccionada] = useState(false);

  const [pagarPrestamo, setPagarPrestamo] = useState(false);
  const [pagarMulta, setPagarMulta] = useState(false);
  const [pagarActividad, setPagarActividad] = useState(false);

  const [montoInteres, setMontoInteres] = useState('0');
  const [montoCapital, setMontoCapital] = useState('0');
  const [montoMulta, setMontoMulta] = useState('3.000');
  const [montoActividad, setMontoActividad] = useState('0');

  const [prestamoActivoSocio, setPrestamoActivoSocio] = useState<Prestamo | null>(null);
  const [mesesInteresCalculados, setMesesInteresCalculados] = useState(0);
  const [totalInteresEstimado, setTotalInteresEstimado] = useState(0);

  const [archivo, setArchivo] = useState<File | null>(null);
  const [cargandoForm, setCargandoForm] = useState(false);
  const [errorForm, setErrorForm] = useState<string | null>(null);
  const [resumenExito, setResumenExito] = useState<{
    nombre: string;
    conceptos: string;
    total: number;
    fecha: string;
    conMulta: boolean;
  } | null>(null);

  // Formulario Admin: Préstamo
  const [socioNuevoPrestamo, setSocioNuevoPrestamo] = useState<number | null>(null);
  const [montoNuevoPrestamo, setMontoNuevoPrestamo] = useState('500.000');
  const [tasaNuevoPrestamo, setTasaNuevoPrestamo] = useState('2.0');
  const [fechaNuevoPrestamo, setFechaNuevoPrestamo] = useState('');

  // Formulario Admin: Actividad
  const [nombreNuevaActividad, setNombreNuevaActividad] = useState('');
  const [valorNuevaActividad, setValorNuevaActividad] = useState('20.000');
  const [premiosNuevaActividad, setPremiosNuevaActividad] = useState('500.000');
  const [fechaNuevaActividad, setFechaNuevaActividad] = useState('');

  // Formulario Admin: Movimiento Manual de Caja
  const [tipoMovManual, setTipoMovManual] = useState<'ingreso_efectivo' | 'gasto_fondo'>('gasto_fondo');
  const [conceptoMovManual, setConceptoMovManual] = useState('');
  const [montoMovManual, setMontoMovManual] = useState('10.000');

  // Estado para Portal de Consulta del Socio
  const [socioConsultaId, setSocioConsultaId] = useState<number | null>(null);

  const [estadoSubidaSocios, setEstadoSubidaSocios] = useState<string | null>(null);

  useEffect(() => {
    cargarSocios();
    cargarActividades();
    cargarMovimientosCaja();
    const hoy = new Date().toISOString().split('T')[0];
    setFechaPago(hoy);
    setFechaNuevoPrestamo(hoy);
    setFechaNuevaActividad(hoy);
  }, []);

  useEffect(() => {
    if (esAdmin) {
      cargarTransacciones();
      cargarPrestamos();
      cargarActividades();
      cargarMovimientosCaja();
    }
  }, [esAdmin]);

  useEffect(() => {
    if (!socioSeleccionado) {
      setPrestamoActivoSocio(null);
      return;
    }
    buscarPrestamoSocio(socioSeleccionado);
  }, [socioSeleccionado]);

  async function cargarSocios() {
    const { data } = await supabase.from('socios').select('*').order('nombre');
    if (data) {
      setSocios(data);
      if (data.length > 0 && !socioParaExtracto) setSocioParaExtracto(data[0].id);
      if (data.length > 0 && !socioConsultaId) setSocioConsultaId(data[0].id);
    }
  }

  async function cargarTransacciones() {
    const { data } = await supabase
      .from('transacciones')
      .select('*, socios(nombre)')
      .order('fecha_reporte', { ascending: false });
    if (data) setTransacciones(data as any);
  }

  async function cargarPrestamos() {
    const { data } = await supabase
      .from('prestamos')
      .select('*, socios(nombre)')
      .order('fecha_desembolso', { ascending: false });
    if (data) setPrestamos(data as any);
  }

  async function cargarActividades() {
    const { data } = await supabase
      .from('actividades')
      .select('*')
      .order('created_at', { ascending: false });
    if (data) setActividades(data);
  }

  async function cargarMovimientosCaja() {
    const { data } = await supabase
      .from('movimientos_caja')
      .select('*')
      .order('created_at', { ascending: false });
    if (data) setMovimientosCaja(data);
  }

  async function buscarPrestamoSocio(idSocio: number) {
    const { data } = await supabase
      .from('prestamos')
      .select('*')
      .eq('socio_id', idSocio)
      .eq('estado', 'activo')
      .maybeSingle();

    if (data) {
      setPrestamoActivoSocio(data);
      const fechaDesembolso = new Date(data.fecha_desembolso);
      const hoy = new Date();
      const difMeses = Math.max(1, (hoy.getFullYear() - fechaDesembolso.getFullYear()) * 12 + (hoy.getMonth() - fechaDesembolso.getMonth()));
      const interesMensual = (Number(data.saldo_capital) * Number(data.tasa_interes_mensual)) / 100;
      setMesesInteresCalculados(difMeses);
      setTotalInteresEstimado(interesMensual * difMeses);
    } else {
      setPrestamoActivoSocio(null);
    }
  }

  function limpiarMonto(valor: string): number {
    const soloNumeros = valor.replace(/[^0-9]/g, '');
    return parseFloat(soloNumeros) || 0;
  }

  const montoAhorroCalculado = (pagarAhorro ? (cuota1Seleccionada ? 50000 : 0) + (cuota2Seleccionada ? 50000 : 0) : 0);
  const actividadActiva = actividades.find((a) => a.activa);

  const totalAConsignar = 
    montoAhorroCalculado +
    (pagarPrestamo ? (limpiarMonto(montoInteres) + limpiarMonto(montoCapital)) : 0) +
    (pagarMulta ? limpiarMonto(montoMulta) : 0) +
    (pagarActividad ? limpiarMonto(montoActividad || (actividadActiva ? String(actividadActiva.valor_cuota) : '0')) : 0);

  async function enviarPago(e: React.FormEvent) {
    e.preventDefault();
    setErrorForm(null);

    if (!socioSeleccionado) {
      setErrorForm('Por favor selecciona tu nombre de la lista.');
      return;
    }
    if (!pagarAhorro && !pagarPrestamo && !pagarMulta && !pagarActividad) {
      setErrorForm('Selecciona al menos un concepto de pago.');
      return;
    }
    if (pagarAhorro && !cuota1Seleccionada && !cuota2Seleccionada) {
      setErrorForm('Debes marcar la Cuota 1, Cuota 2 o ambas.');
      return;
    }
    if (totalAConsignar <= 0) {
      setErrorForm('El total a pagar debe ser mayor a $0.');
      return;
    }
    if (!archivo) {
      setErrorForm('Por favor adjunta el comprobante de pago.');
      return;
    }

    setCargandoForm(true);

    try {
      const extension = archivo.name.split('.').pop();
      const nombreArchivo = `${socioSeleccionado}-${Date.now()}.${extension}`;
      const { error: uploadError } = await supabase.storage
        .from('comprobantes')
        .upload(nombreArchivo, archivo);

      if (uploadError) throw uploadError;

      const { data: urlData } = supabase.storage
        .from('comprobantes')
        .getPublicUrl(nombreArchivo);

      const fechaFinal = new Date(`${fechaPago}T12:00:00Z`).toISOString();

      const conceptosList: string[] = [];
      if (pagarAhorro) {
        if (cuota1Seleccionada && cuota2Seleccionada) conceptosList.push('Ahorro (Mes Completo $100.000)');
        else if (cuota1Seleccionada) conceptosList.push('Ahorro (Cuota 1 $50.000)');
        else if (cuota2Seleccionada) conceptosList.push('Ahorro (Cuota 2 $50.000)');
      }
      if (pagarPrestamo) conceptosList.push('Abono Préstamo');
      if (pagarMulta) conceptosList.push('Multa');
      if (pagarActividad) conceptosList.push(`Actividad (${actividadActiva?.nombre || 'Rifa'})`);

      const { error: insertError } = await supabase.from('transacciones').insert({
        socio_id: socioSeleccionado,
        monto_reportado: totalAConsignar,
        concepto: conceptosList.join(', '),
        pago_cuota_1: pagarAhorro && cuota1Seleccionada,
        pago_cuota_2: pagarAhorro && cuota2Seleccionada,
        desglose_ahorro: montoAhorroCalculado,
        desglose_prestamo_capital: pagarPrestamo ? limpiarMonto(montoCapital) : 0,
        desglose_prestamo_interes: pagarPrestamo ? limpiarMonto(montoInteres) : 0,
        desglose_multa: pagarMulta ? limpiarMonto(montoMulta) : 0,
        desglose_actividad: pagarActividad ? limpiarMonto(montoActividad || String(actividadActiva?.valor_cuota || 0)) : 0,
        comprobante_url: urlData.publicUrl,
        fecha_reporte: fechaFinal,
        estado: 'pendiente',
      });

      if (insertError) throw insertError;

      const socioObj = socios.find((s) => s.id === socioSeleccionado);
      setResumenExito({
        nombre: socioObj?.nombre || 'Socio',
        conceptos: conceptosList.join(', '),
        total: totalAConsignar,
        fecha: fechaPago,
        conMulta: esFechaExtemporanea(fechaPago),
      });

      setArchivo(null);
    } catch (err: any) {
      setErrorForm(err.message || 'Error al procesar el reporte.');
    } finally {
      setCargandoForm(false);
    }
  }

  async function actualizarEstado(t: Transaccion, nuevoEstado: 'aprobado' | 'rechazado') {
    const { error } = await supabase
      .from('transacciones')
      .update({ estado: nuevoEstado, fecha_aprobacion: new Date().toISOString() })
      .eq('id', t.id);

    if (!error) {
      if (nuevoEstado === 'aprobado' && t.desglose_prestamo_capital > 0) {
        const prestamo = prestamos.find((p) => p.socio_id === t.socio_id && p.estado === 'activo');
        if (prestamo) {
          const nuevoSaldo = Math.max(0, Number(prestamo.saldo_capital) - Number(t.desglose_prestamo_capital));
          await supabase
            .from('prestamos')
            .update({ 
              saldo_capital: nuevoSaldo,
              estado: nuevoSaldo === 0 ? 'pagado' : 'activo'
            })
            .eq('id', prestamo.id);
          cargarPrestamos();
        }
      }

      setTransacciones((prev) =>
        prev.map((item) => (item.id === t.id ? { ...item, estado: nuevoEstado } : item))
      );
    }
  }

  async function crearPrestamo(e: React.FormEvent) {
    e.preventDefault();
    if (!socioNuevoPrestamo) return;

    const monto = limpiarMonto(montoNuevoPrestamo);
    const tasa = parseFloat(tasaNuevoPrestamo) || 2.0;

    const { error } = await supabase.from('prestamos').insert({
      socio_id: socioNuevoPrestamo,
      monto_prestado: monto,
      saldo_capital: monto,
      tasa_interes_mensual: tasa,
      fecha_desembolso: fechaNuevoPrestamo,
      estado: 'activo'
    });

    if (!error) {
      alert('Préstamo registrado exitosamente');
      setSocioNuevoPrestamo(null);
      cargarPrestamos();
    } else {
      alert('Error: ' + error.message);
    }
  }

  async function liquidarPrestamo(idPrestamo: number) {
    if (!confirm('¿Deseas marcar este préstamo como Paz y Salvo? Pasará a estado pagado.')) return;
    const { error } = await supabase
      .from('prestamos')
      .update({ estado: 'pagado', saldo_capital: 0 })
      .eq('id', idPrestamo);

    if (!error) {
      alert('Préstamo liquidado correctamente.');
      cargarPrestamos();
    }
  }

  async function crearActividad(e: React.FormEvent) {
    e.preventDefault();
    if (!nombreNuevaActividad.trim()) return;

    const { error } = await supabase.from('actividades').insert({
      nombre: nombreNuevaActividad.trim(),
      valor_cuota: limpiarMonto(valorNuevaActividad),
      costo_premios: limpiarMonto(premiosNuevaActividad),
      fecha_limite: fechaNuevaActividad,
      activa: true
    });

    if (!error) {
      alert('Actividad creada exitosamente');
      setNombreNuevaActividad('');
      cargarActividades();
    } else {
      alert('Error: ' + error.message);
    }
  }

  async function cambiarEstadoActividad(id: number, estadoActual: boolean) {
    await supabase.from('actividades').update({ activa: !estadoActual }).eq('id', id);
    cargarActividades();
  }

  async function registrarMovimientoCaja(e: React.FormEvent) {
    e.preventDefault();
    if (!conceptoMovManual.trim()) return;

    const valor = limpiarMonto(montoMovManual);
    if (valor <= 0) return;

    const { error } = await supabase.from('movimientos_caja').insert({
      tipo: tipoMovManual,
      concepto: conceptoMovManual.trim(),
      monto: valor,
      fecha: new Date().toISOString().split('T')[0]
    });

    if (!error) {
      alert('Movimiento registrado en caja.');
      setConceptoMovManual('');
      cargarMovimientosCaja();
    } else {
      alert('Error: ' + error.message);
    }
  }

  async function ejecutarBorradoTotal() {
    if (pinBorrado !== PIN_CORRECTO) {
      alert('Clave incorrecta. Acción cancelada.');
      return;
    }

    setCargandoBorrado(true);
    try {
      await supabase.from('transacciones').delete().neq('id', 0);
      await supabase.from('prestamos').delete().neq('id', 0);
      await supabase.from('actividades').delete().neq('id', 0);
      await supabase.from('movimientos_caja').delete().neq('id', 0);

      alert('¡Todo limpio! Se han borrado transacciones, préstamos, actividades y movimientos de caja.');
      setMostrarModalBorrado(false);
      setPinBorrado('');
      cargarTransacciones();
      cargarPrestamos();
      cargarActividades();
      cargarMovimientosCaja();
    } catch (err: any) {
      alert('Error al borrar: ' + err.message);
    } finally {
      setCargandoBorrado(false);
    }
  }

  async function procesarExcelSocios(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    setEstadoSubidaSocios('Leyendo archivo Excel...');
    const reader = new FileReader();

    reader.onload = async (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(buffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const matriz: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

        let colNombre = 0;
        let inicioFila = 0;

        for (let f = 0; f < Math.min(matriz.length, 5); f++) {
          const fila = matriz[f] || [];
          for (let c = 0; c < fila.length; c++) {
            const txt = String(fila[c] || '').toLowerCase().trim();
            if (txt.includes('nombre') || txt.includes('socio')) {
              colNombre = c;
              inicioFila = f + 1;
              break;
            }
          }
        }

        const lista = [];
        for (let i = inicioFila; i < matriz.length; i++) {
          const val = String(matriz[i]?.[colNombre] || '').trim();
          if (val.length > 1 && !val.toLowerCase().includes('nombre')) {
            lista.push({ nombre: val });
          }
        }

        const { error } = await supabase.from('socios').upsert(lista, { onConflict: 'nombre' });
        if (error) throw error;
        setEstadoSubidaSocios(`¡Éxito! Se cargaron ${lista.length} socios.`);
        cargarSocios();
      } catch (err: any) {
        setEstadoSubidaSocios('Error: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function verificarPin() {
    if (pinIngresado === PIN_CORRECTO) {
      setEsAdmin(true);
      setMostrarModalPin(false);
      setPinIngresado('');
    } else {
      alert('Clave incorrecta');
    }
  }

  // --- CÁLCULOS DEL BALANCE GENERAL ---
  const transaccionesAprobadas = transacciones.filter((t) => t.estado === 'aprobado');
  
  const totalBancoRecaudado = transaccionesAprobadas.reduce((acc, t) => acc + Number(t.monto_reportado), 0);
  const totalFondoAhorros = transaccionesAprobadas.reduce((acc, t) => acc + Number(t.desglose_ahorro), 0);
  const totalInteresesCobrados = transaccionesAprobadas.reduce((acc, t) => acc + Number(t.desglose_prestamo_interes), 0);
  const totalMultasCobradas = transaccionesAprobadas.reduce((acc, t) => acc + Number(t.desglose_multa), 0);
  const totalActividadesBruto = transaccionesAprobadas.reduce((acc, t) => acc + Number(t.desglose_actividad), 0);
  
  const totalPremiosEntregados = actividades.reduce((acc, a) => acc + Number(a.costo_premios || 0), 0);
  const gananciaNetaActividades = Math.max(0, totalActividadesBruto - totalPremiosEntregados);

  // Movimientos de caja manuales
  const totalIngresosEfectivoManual = movimientosCaja
    .filter((m) => m.tipo === 'ingreso_efectivo')
    .reduce((acc, m) => acc + Number(m.monto), 0);

  const totalGastosFondoManual = movimientosCaja
    .filter((m) => m.tipo === 'gasto_fondo')
    .reduce((acc, m) => acc + Number(m.monto), 0);

  const capitalPrestadoActivo = prestamos
    .filter((p) => p.estado === 'activo')
    .reduce((acc, p) => acc + Number(p.saldo_capital), 0);

  const dineroDisponibleParaPrestar = Math.max(
    0, 
    (totalBancoRecaudado + totalIngresosEfectivoManual) - capitalPrestadoActivo - totalGastosFondoManual
  );

  const rendimientosTotales = totalInteresesCobrados + totalMultasCobradas + gananciaNetaActividades - totalGastosFondoManual;

  // --- CÁLCULOS DE ATRASADOS Y MULTAS ---
  const prefijoMes = `${anioSeleccionado}-${mesSeleccionado}`;
  const nombreMes = MESES.find((m) => m.valor === mesSeleccionado)?.nombre || '';

  const sociosDetalleMes = socios.map((s) => {
    const pagosMes = transaccionesAprobadas.filter(
      (t) => t.socio_id === s.id && t.fecha_reporte.startsWith(prefijoMes)
    );
    const pagoCuota1 = pagosMes.some((t) => t.pago_cuota_1);
    const pagoCuota2 = pagosMes.some((t) => t.pago_cuota_2);
    
    let estado = 'al_dia';
    let textoAtraso = 'Al día';
    if (!pagoCuota1 && !pagoCuota2) {
      estado = 'debe_todo';
      textoAtraso = 'Debe Mes Completo ($100.000)';
    } else if (!pagoCuota1) {
      estado = 'debe_c1';
      textoAtraso = 'Debe Cuota 1 ($50.000)';
    } else if (!pagoCuota2) {
      estado = 'debe_c2';
      textoAtraso = 'Debe Cuota 2 ($50.000)';
    }

    const cuotasExtemporaneas = transaccionesAprobadas.filter(
      (t) => t.socio_id === s.id && esFechaExtemporanea(t.fecha_reporte) && (t.desglose_ahorro > 0 || t.desglose_actividad > 0)
    );
    const totalMultasCausadas = cuotasExtemporaneas.length * 3000;

    const totalMultasAbonadas = transaccionesAprobadas
      .filter((t) => t.socio_id === s.id)
      .reduce((acc, t) => acc + Number(t.desglose_multa || 0), 0);

    const multasPendientesValor = Math.max(0, totalMultasCausadas - totalMultasAbonadas);

    return { 
      ...s, 
      pagoCuota1, 
      pagoCuota2, 
      estado, 
      textoAtraso, 
      multasPendientesValor,
      totalMultasAbonadas
    };
  });

  // Filtros Comprobantes
  const transaccionesFiltradas = transacciones.filter((t) => {
    if (filtroComprobanteSocio !== 'todos' && String(t.socio_id) !== filtroComprobanteSocio) return false;
    if (filtroComprobanteMes !== 'todos') {
      const mesTx = t.fecha_reporte.slice(5, 7);
      if (mesTx !== filtroComprobanteMes) return false;
    }
    return true;
  });

  // Cálculos para Portal de Consulta del Socio
  const socioConsultaObj = socios.find((s) => s.id === socioConsultaId);
  const pagosConsulta = transacciones.filter((t) => t.socio_id === socioConsultaId);
  const totalAhorradoConsulta = pagosConsulta
    .filter((t) => t.estado === 'aprobado')
    .reduce((acc, t) => acc + Number(t.desglose_ahorro), 0);

  const prestamoConsulta = prestamos.find((p) => p.socio_id === socioConsultaId && p.estado === 'activo');
  const multasPagadasConsulta = pagosConsulta
    .filter((t) => t.estado === 'aprobado')
    .reduce((acc, t) => acc + Number(t.desglose_multa), 0);

  const multasCausadasConsulta = pagosConsulta
    .filter((t) => t.estado === 'aprobado' && esFechaExtemporanea(t.fecha_reporte) && t.desglose_ahorro > 0)
    .length * 3000;

  const saldoMultasConsulta = Math.max(0, multasCausadasConsulta - multasPagadasConsulta);

  // Exportar Excel
  function exportarBalanceCompleto() {
    const wb = XLSX.utils.book_new();

    const resumen = [
      { Concepto: 'Total Dinero en Banco (Recaudado Bre-B)', Valor: totalBancoRecaudado },
      { Concepto: 'Ingresos Efectivo Registrados Manual', Valor: totalIngresosEfectivoManual },
      { Concepto: 'Gastos de Operación Registrados Manual', Valor: totalGastosFondoManual },
      { Concepto: 'Fondo Ahorros de Socios', Valor: totalFondoAhorros },
      { Concepto: 'Capital Activo en Préstamos', Valor: capitalPrestadoActivo },
      { Concepto: 'Disponible Real en Caja', Valor: dineroDisponibleParaPrestar },
      { Concepto: 'Ganancia Intereses Cobrados', Valor: totalInteresesCobrados },
      { Concepto: 'Recaudo Multas', Valor: totalMultasCobradas },
      { Concepto: 'Ganancia Neta Actividades', Valor: gananciaNetaActividades },
      { Concepto: 'Rendimientos Líquidos para Dividendos', Valor: rendimientosTotales }
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), 'Resumen_Balance');

    const transaccionesFormato = transaccionesAprobadas.map((t) => ({
      Fecha: new Date(t.fecha_reporte).toLocaleDateString(),
      Socio: t.socios?.nombre || 'Desconocido',
      Total_Consignado: t.monto_reportado,
      Ahorro: t.desglose_ahorro,
      Interes_Prestamo: t.desglose_prestamo_interes,
      Abono_Capital: t.desglose_prestamo_capital,
      Multa: t.desglose_multa,
      Actividad: t.desglose_actividad,
      Concepto: t.concepto
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(transaccionesFormato), 'Transacciones');

    XLSX.writeFile(wb, `Balance_General_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function exportarRecaudoMes() {
    const pagosMes = transaccionesAprobadas.filter((t) => t.fecha_reporte.startsWith(prefijoMes));
    if (pagosMes.length === 0) {
      alert(`No hay registros aprobados para ${nombreMes} ${anioSeleccionado}.`);
      return;
    }

    const datos = pagosMes.map((t) => ({
      Fecha: new Date(t.fecha_reporte).toLocaleDateString(),
      Socio: t.socios?.nombre || 'Desconocido',
      Total: t.monto_reportado,
      Ahorro: t.desglose_ahorro,
      Interes: t.desglose_prestamo_interes,
      Capital: t.desglose_prestamo_capital,
      Multa: t.desglose_multa,
      Actividad: t.desglose_actividad,
      Detalle: t.concepto
    }));

    const ws = XLSX.utils.json_to_sheet(datos);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Recaudo_${nombreMes}`);
    XLSX.writeFile(wb, `Recaudo_${nombreMes}_${anioSeleccionado}.xlsx`);
  }

  function exportarExtractoSocio() {
    if (!socioParaExtracto) return;
    const socioObj = socios.find((s) => s.id === socioParaExtracto);
    if (!socioObj) return;

    const pagosSocio = transaccionesAprobadas.filter((t) => t.socio_id === socioParaExtracto);
    const prestamoSocio = prestamos.find((p) => p.socio_id === socioParaExtracto && p.estado === 'activo');

    const totalAhorrado = pagosSocio.reduce((acc, t) => acc + Number(t.desglose_ahorro), 0);
    const totalInteresesPagados = pagosSocio.reduce((acc, t) => acc + Number(t.desglose_prestamo_interes), 0);
    const totalCapitalAbonado = pagosSocio.reduce((acc, t) => acc + Number(t.desglose_prestamo_capital), 0);

    const encabezado = [
      { Campo: 'Socio', Detalle: socioObj.nombre },
      { Campo: 'Total Ahorro Acumulado', Detalle: totalAhorrado },
      { Campo: 'Saldo Capital Pendiente', Detalle: prestamoSocio ? prestamoSocio.saldo_capital : 0 },
      { Campo: 'Total Intereses Pagados', Detalle: totalInteresesPagados },
      { Campo: 'Total Capital Amortizado', Detalle: totalCapitalAbonado }
    ];

    const detallePagos = pagosSocio.map((t) => ({
      Fecha: new Date(t.fecha_reporte).toLocaleDateString(),
      Total: t.monto_reportado,
      Ahorro: t.desglose_ahorro,
      Interes: t.desglose_prestamo_interes,
      Capital: t.desglose_prestamo_capital,
      Multa: t.desglose_multa,
      Actividad: t.desglose_actividad,
      Concepto: t.concepto
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(encabezado), 'Resumen');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detallePagos), 'Historial');
    XLSX.writeFile(wb, `Extracto_${socioObj.nombre.replace(/\s+/g, '_')}.xlsx`);
  }

  return (
    <main className="min-h-screen bg-slate-900 text-slate-100 p-4 md:p-8">
      {/* Cabecera superior */}
      <div className="max-w-6xl mx-auto flex justify-between items-center mb-6">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-emerald-500/20 text-emerald-400 rounded-xl">
            <Wallet className="w-6 h-6" />
          </div>
          <div>
            <h1 className="font-bold text-lg text-white">Fondo de Ahorro</h1>
            <p className="text-xs text-slate-400">Sistema Integral 2027</p>
          </div>
        </div>

        {esAdmin ? (
          <div className="flex items-center gap-2">
            <button
              onClick={() => setMostrarModalBorrado(true)}
              className="flex items-center gap-1.5 bg-red-600/20 hover:bg-red-600 text-red-300 hover:text-white px-3 py-2 rounded-xl border border-red-500/30 text-xs font-semibold transition"
            >
              <Trash2 className="w-4 h-4" /> Reiniciar Pruebas
            </button>
            <button
              onClick={() => setEsAdmin(false)}
              className="flex items-center gap-2 bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2 rounded-xl border border-slate-700 text-sm font-medium transition"
            >
              <LogOut className="w-4 h-4 text-red-400" /> Salir de Admin
            </button>
          </div>
        ) : (
          <div className="flex items-center gap-2">
            {/* Selector de modo usuario */}
            <div className="bg-slate-800 p-1 rounded-xl border border-slate-700 flex text-xs">
              <button
                onClick={() => setModoPublico('reportar')}
                className={`px-3 py-1.5 rounded-lg transition ${modoPublico === 'reportar' ? 'bg-emerald-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}
              >
                Reportar Pago
              </button>
              <button
                onClick={() => { setModoPublico('consultar'); cargarTransacciones(); cargarPrestamos(); }}
                className={`px-3 py-1.5 rounded-lg transition ${modoPublico === 'consultar' ? 'bg-blue-600 text-white font-semibold' : 'text-slate-400 hover:text-white'}`}
              >
                Consultar Mi Estado
              </button>
            </div>

            <button
              onClick={() => setMostrarModalPin(true)}
              className="flex items-center gap-2 bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2 rounded-xl border border-slate-700 text-xs font-medium transition shadow-md"
            >
              <Lock className="w-3.5 h-3.5 text-emerald-400" /> Admin
            </button>
          </div>
        )}
      </div>

      {/* Modal PIN */}
      {mostrarModalPin && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-800 border border-slate-700 rounded-2xl p-6 max-w-xs w-full shadow-2xl space-y-4 text-center">
            <ShieldCheck className="w-8 h-8 text-emerald-400 mx-auto" />
            <h2 className="text-lg font-bold text-white">Clave de Administrador</h2>
            <input
              type="password"
              placeholder="1234"
              value={pinIngresado}
              onChange={(e) => setPinIngresado(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && verificarPin()}
              className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-center text-white text-lg tracking-widest focus:outline-none focus:border-emerald-500"
              autoFocus
            />
            <div className="flex gap-2">
              <button onClick={() => setMostrarModalPin(false)} className="w-1/2 bg-slate-700 py-2 rounded-lg text-sm">Cancelar</button>
              <button onClick={verificarPin} className="w-1/2 bg-emerald-600 hover:bg-emerald-500 py-2 rounded-lg text-sm font-semibold">Entrar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Borrado Seguro */}
      {mostrarModalBorrado && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-800 border border-red-500/50 rounded-2xl p-6 max-w-sm w-full shadow-2xl space-y-4 text-center">
            <AlertTriangle className="w-8 h-8 text-red-400 mx-auto" />
            <h2 className="text-lg font-bold text-white">¿Reiniciar Datos para 2027?</h2>
            <p className="text-xs text-slate-300">
              Se eliminarán transacciones, préstamos, actividades y movimientos de caja. La lista de socios se mantiene.
            </p>
            <input
              type="password"
              placeholder="1234"
              value={pinBorrado}
              onChange={(e) => setPinBorrado(e.target.value)}
              className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2 text-center text-white tracking-widest"
              autoFocus
            />
            <div className="flex gap-2 pt-2">
              <button onClick={() => { setMostrarModalBorrado(false); setPinBorrado(''); }} className="w-1/2 bg-slate-700 py-2 rounded-lg text-sm">Cancelar</button>
              <button onClick={ejecutarBorradoTotal} disabled={cargandoBorrado || !pinBorrado} className="w-1/2 bg-red-600 py-2 rounded-lg text-sm font-semibold">Confirmar</button>
            </div>
          </div>
        </div>
      )}

      {/* VISTA ADMIN */}
      {esAdmin ? (
        <div className="max-w-6xl mx-auto space-y-6">
          <div className="flex flex-wrap gap-2 border-b border-slate-700 pb-3">
            <button
              onClick={() => setPestanaAdmin('balance')}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 ${
                pestanaAdmin === 'balance' ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <PieChart className="w-4 h-4" /> Balance General
            </button>
            <button
              onClick={() => setPestanaAdmin('comprobantes')}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 ${
                pestanaAdmin === 'comprobantes' ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <FileSpreadsheet className="w-4 h-4" /> Comprobantes Bre-B
            </button>
            <button
              onClick={() => setPestanaAdmin('prestamos')}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 ${
                pestanaAdmin === 'prestamos' ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <CreditCard className="w-4 h-4" /> Préstamos
            </button>
            <button
              onClick={() => setPestanaAdmin('actividades')}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 ${
                pestanaAdmin === 'actividades' ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <Tag className="w-4 h-4" /> Actividades / Rifas
            </button>
            <button
              onClick={() => setPestanaAdmin('atrasados')}
              className={`px-4 py-2 rounded-xl text-sm font-medium transition flex items-center gap-2 ${
                pestanaAdmin === 'atrasados' ? 'bg-emerald-600 text-white' : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
              }`}
            >
              <AlertCircle className="w-4 h-4" /> Control de Atrasados y Multas
            </button>
          </div>

          {/* BALANCE GENERAL */}
          {pestanaAdmin === 'balance' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="bg-slate-800 border border-emerald-500/40 p-5 rounded-2xl space-y-2">
                  <div className="flex justify-between items-center text-slate-400 text-xs">
                    <span>DINERO TOTAL EN BANCO</span>
                    <Landmark className="w-5 h-5 text-emerald-400" />
                  </div>
                  <div className="text-2xl font-black text-white">
                    ${formatearMoneda(totalBancoRecaudado)} <span className="text-xs text-slate-400 font-normal">COP</span>
                  </div>
                  <p className="text-xs text-slate-400">Total acumulado de reportes aprobados</p>
                </div>

                <div className="bg-slate-800 border border-blue-500/40 p-5 rounded-2xl space-y-2">
                  <div className="flex justify-between items-center text-slate-400 text-xs">
                    <span>DISPONIBLE REAL EN CAJA</span>
                    <Coins className="w-5 h-5 text-blue-400" />
                  </div>
                  <div className="text-2xl font-black text-blue-400">
                    ${formatearMoneda(dineroDisponibleParaPrestar)} <span className="text-xs text-slate-400 font-normal">COP</span>
                  </div>
                  <p className="text-xs text-slate-400">Ajustado con préstamos y movimientos manuales</p>
                </div>

                <div className="bg-slate-800 border border-amber-500/40 p-5 rounded-2xl space-y-2">
                  <div className="flex justify-between items-center text-slate-400 text-xs">
                    <span>CAPITAL EN PRÉSTAMOS</span>
                    <CreditCard className="w-5 h-5 text-amber-400" />
                  </div>
                  <div className="text-2xl font-black text-amber-400">
                    ${formatearMoneda(capitalPrestadoActivo)} <span className="text-xs text-slate-400 font-normal">COP</span>
                  </div>
                  <p className="text-xs text-slate-400">Dinero colocado en manos de socios</p>
                </div>
              </div>

              {/* MEJORA 3: MOVIMIENTOS MANUALES DE CAJA (INGRESOS EFECTIVO / GASTOS) */}
              <div className="bg-slate-800 p-5 rounded-2xl border border-slate-700 space-y-4">
                <h3 className="font-bold text-white text-sm flex items-center gap-2">
                  <Coins className="w-4 h-4 text-amber-400" /> Registro Manual de Caja (Efectivo o Gastos de Papelería)
                </h3>
                <form onSubmit={registrarMovimientoCaja} className="grid grid-cols-1 md:grid-cols-4 gap-3">
                  <select
                    value={tipoMovManual}
                    onChange={(e: any) => setTipoMovManual(e.target.value)}
                    className="bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs"
                  >
                    <option value="gasto_fondo">🔴 Gasto de Operación / Papelería</option>
                    <option value="ingreso_efectivo">🟢 Ingreso Recibido en Efectivo</option>
                  </select>
                  <input
                    type="text"
                    placeholder="Descripción (ej: Fotocopias de actas)"
                    value={conceptoMovManual}
                    onChange={(e) => setConceptoMovManual(e.target.value)}
                    className="bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs"
                    required
                  />
                  <input
                    type="text"
                    placeholder="Monto (ej: 15.000)"
                    value={montoMovManual}
                    onChange={(e) => setMontoMovManual(e.target.value)}
                    className="bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs font-semibold"
                    required
                  />
                  <button type="submit" className="bg-slate-700 hover:bg-slate-600 text-white font-semibold rounded-lg text-xs py-2 transition border border-slate-600">
                    Guardar Movimiento
                  </button>
                </form>
              </div>

              {/* Desglose de Fondos */}
              <div className="bg-slate-800 p-6 rounded-2xl border border-slate-700 space-y-4">
                <h3 className="font-bold text-white text-base">Distribución del Capital Recaudado</h3>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                  <div className="p-3 bg-slate-700/40 rounded-xl border border-slate-600">
                    <span className="text-xs text-slate-400 block mb-1">Ahorros de Socios</span>
                    <span className="font-bold text-white text-base">${formatearMoneda(totalFondoAhorros)}</span>
                  </div>
                  <div className="p-3 bg-slate-700/40 rounded-xl border border-slate-600">
                    <span className="text-xs text-emerald-400 block mb-1">Ganancia Intereses</span>
                    <span className="font-bold text-emerald-400 text-base">${formatearMoneda(totalInteresesCobrados)}</span>
                  </div>
                  <div className="p-3 bg-slate-700/40 rounded-xl border border-slate-600">
                    <span className="text-xs text-red-400 block mb-1">Recaudo Multas</span>
                    <span className="font-bold text-red-400 text-base">${formatearMoneda(totalMultasCobradas)}</span>
                  </div>
                  <div className="p-3 bg-slate-700/40 rounded-xl border border-slate-600">
                    <span className="text-xs text-purple-400 block mb-1">Ganancia Neta Actividades</span>
                    <span className="font-bold text-purple-400 text-base">${formatearMoneda(gananciaNetaActividades)}</span>
                    <span className="text-[10px] text-slate-400 block mt-0.5">Descontando premios</span>
                  </div>
                </div>

                <div className="p-4 bg-emerald-500/10 border border-emerald-500/30 rounded-xl flex justify-between items-center text-xs md:text-sm">
                  <span className="text-emerald-300 font-medium">Rendimientos Líquidos del Fondo (Para dividendos):</span>
                  <span className="font-extrabold text-emerald-400 text-base">${formatearMoneda(rendimientosTotales)} COP</span>
                </div>
              </div>

              {/* SECCIÓN DE DESCARGAS EN EXCEL */}
              <div className="bg-slate-800 p-6 rounded-2xl border border-slate-700 space-y-4">
                <h3 className="font-bold text-white text-base flex items-center gap-2">
                  <Download className="w-5 h-5 text-emerald-400" /> Exportar Reportes Contables en Excel
                </h3>
                
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="p-4 bg-slate-700/30 rounded-xl border border-slate-600 flex flex-col justify-between space-y-3">
                    <div>
                      <h4 className="font-bold text-white text-sm flex items-center gap-1.5">
                        <FileSpreadsheet className="w-4 h-4 text-emerald-400" /> Balance General Anual
                      </h4>
                      <p className="text-xs text-slate-400 mt-1">Consolidado con resumen contable y transacciones.</p>
                    </div>
                    <button
                      onClick={exportarBalanceCompleto}
                      className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-2 rounded-lg text-xs flex items-center justify-center gap-1.5 transition"
                    >
                      <Download className="w-3.5 h-3.5" /> Descargar Balance Completo
                    </button>
                  </div>

                  <div className="p-4 bg-slate-700/30 rounded-xl border border-slate-600 flex flex-col justify-between space-y-3">
                    <div>
                      <h4 className="font-bold text-white text-sm flex items-center gap-1.5">
                        <FileText className="w-4 h-4 text-blue-400" /> Recaudo Mensual Específico
                      </h4>
                      <p className="text-xs text-slate-400 mt-1">Exporta todos los ingresos detallados de: <strong>{nombreMes} {anioSeleccionado}</strong>.</p>
                    </div>
                    <button
                      onClick={exportarRecaudoMes}
                      className="w-full bg-blue-600 hover:bg-blue-500 text-white font-semibold py-2 rounded-lg text-xs flex items-center justify-center gap-1.5 transition"
                    >
                      <Download className="w-3.5 h-3.5" /> Descargar Mes ({nombreMes})
                    </button>
                  </div>

                  <div className="p-4 bg-slate-700/30 rounded-xl border border-slate-600 flex flex-col justify-between space-y-3">
                    <div>
                      <h4 className="font-bold text-white text-sm flex items-center gap-1.5">
                        <UserCheck className="w-4 h-4 text-purple-400" /> Extracto Individual de Socio
                      </h4>
                      <p className="text-xs text-slate-400 mt-1">Informe individual listo para enviar por WhatsApp:</p>
                      <select
                        value={socioParaExtracto || ''}
                        onChange={(e) => setSocioParaExtracto(Number(e.target.value))}
                        className="w-full mt-2 bg-slate-800 border border-slate-600 rounded-lg p-1.5 text-white text-xs"
                      >
                        {socios.map((s) => (
                          <option key={s.id} value={s.id}>{s.nombre}</option>
                        ))}
                      </select>
                    </div>
                    <button
                      onClick={exportarExtractoSocio}
                      className="w-full bg-purple-600 hover:bg-purple-500 text-white font-semibold py-2 rounded-lg text-xs flex items-center justify-center gap-1.5 transition"
                    >
                      <Download className="w-3.5 h-3.5" /> Descargar Extracto Socio
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* COMPROBANTES: CON DETECCIÓN DE DUPLICADOS (MEJORA 2) */}
          {pestanaAdmin === 'comprobantes' && (
            <div className="space-y-4">
              <div className="bg-slate-800 p-4 rounded-xl border border-slate-700 space-y-3">
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-2">
                  <span className="font-semibold text-white">Bandeja de Conciliación Bancaria</span>
                  <div className="flex gap-2">
                    <label className="bg-blue-600 hover:bg-blue-500 text-white px-3 py-1.5 rounded-lg cursor-pointer text-xs font-medium flex items-center gap-1">
                      <Users className="w-3.5 h-3.5" /> Cargar Socios Excel
                      <input type="file" accept=".xlsx, .xls" onChange={procesarExcelSocios} className="hidden" />
                    </label>
                    <button onClick={cargarTransacciones} className="p-1.5 bg-slate-700 rounded-lg text-slate-300 hover:text-white">
                      <RefreshCw className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-3 pt-2 border-t border-slate-700/60">
                  <div className="flex items-center gap-2 text-xs text-slate-300">
                    <Filter className="w-3.5 h-3.5 text-emerald-400" />
                    <span>Filtrar por Socio:</span>
                    <select
                      value={filtroComprobanteSocio}
                      onChange={(e) => setFiltroComprobanteSocio(e.target.value)}
                      className="bg-slate-700 border border-slate-600 rounded-lg px-2.5 py-1 text-white text-xs"
                    >
                      <option value="todos">Todos los socios</option>
                      {socios.map((s) => (
                        <option key={s.id} value={String(s.id)}>{s.nombre}</option>
                      ))}
                    </select>
                  </div>

                  <div className="flex items-center gap-2 text-xs text-slate-300">
                    <span>Filtrar por Mes:</span>
                    <select
                      value={filtroComprobanteMes}
                      onChange={(e) => setFiltroComprobanteMes(e.target.value)}
                      className="bg-slate-700 border border-slate-600 rounded-lg px-2.5 py-1 text-white text-xs"
                    >
                      <option value="todos">Todos los meses</option>
                      {MESES.map((m) => (
                        <option key={m.valor} value={m.valor}>{m.nombre}</option>
                      ))}
                    </select>
                  </div>

                  {(filtroComprobanteSocio !== 'todos' || filtroComprobanteMes !== 'todos') && (
                    <button
                      onClick={() => { setFiltroComprobanteSocio('todos'); setFiltroComprobanteMes('todos'); }}
                      className="text-xs text-red-400 hover:underline"
                    >
                      Limpiar filtros
                    </button>
                  )}
                </div>
              </div>

              {estadoSubidaSocios && <p className="text-xs text-blue-400">{estadoSubidaSocios}</p>}

              <div className="bg-slate-800 rounded-2xl border border-slate-700 overflow-x-auto">
                <table className="w-full text-left text-xs md:text-sm text-slate-300">
                  <thead className="bg-slate-700/50 text-slate-400 uppercase text-xs">
                    <tr>
                      <th className="p-3">Fecha</th>
                      <th className="p-3">Socio</th>
                      <th className="p-3">Total Bre-B</th>
                      <th className="p-3">Concepto Detallado</th>
                      <th className="p-3">Auditoría Fecha</th>
                      <th className="p-3">Comprobante</th>
                      <th className="p-3">Estado</th>
                      <th className="p-3 text-center">Acción</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-700/50">
                    {transaccionesFiltradas.length === 0 ? (
                      <tr><td colSpan={8} className="p-8 text-center text-slate-400">No hay comprobantes para los filtros seleccionados.</td></tr>
                    ) : (
                      transaccionesFiltradas.map((t) => {
                        const esFueraDeTiempo = esFechaExtemporanea(t.fecha_reporte);
                        const faltoMulta = esFueraDeTiempo && Number(t.desglose_multa) === 0;

                        // Detección de duplicado: mismo socio, mismo monto y misma fecha ya aprobada
                        const esPosibleDuplicado = transaccionesAprobadas.some(
                          (ap) => ap.id !== t.id && ap.socio_id === t.socio_id && Number(ap.monto_reportado) === Number(t.monto_reportado) && ap.fecha_reporte.slice(0, 10) === t.fecha_reporte.slice(0, 10)
                        );

                        return (
                          <tr key={t.id} className="hover:bg-slate-700/30">
                            <td className="p-3">{new Date(t.fecha_reporte).toLocaleDateString()}</td>
                            <td className="p-3 font-medium text-white">
                              {t.socios?.nombre || 'Desconocido'}
                              {esPosibleDuplicado && (
                                <span className="block text-[10px] text-amber-400 font-bold">⚠️ Posible pago duplicado</span>
                              )}
                            </td>
                            <td className="p-3 font-bold text-emerald-400">${formatearMoneda(Number(t.monto_reportado))}</td>
                            <td className="p-3 text-xs space-y-0.5">
                              <div className="font-semibold text-slate-200">{t.concepto}</div>
                              {t.desglose_prestamo_interes > 0 && <div className="text-amber-300">• Interés: ${formatearMoneda(Number(t.desglose_prestamo_interes))}</div>}
                              {t.desglose_prestamo_capital > 0 && <div className="text-blue-300">• Abono Capital: ${formatearMoneda(Number(t.desglose_prestamo_capital))}</div>}
                              {t.desglose_multa > 0 && <div className="text-red-300">• Multa: ${formatearMoneda(Number(t.desglose_multa))}</div>}
                              {t.desglose_actividad > 0 && <div className="text-purple-300">• Actividad: ${formatearMoneda(Number(t.desglose_actividad))}</div>}
                            </td>
                            <td className="p-3 text-xs">
                              {faltoMulta ? (
                                <span className="px-2 py-0.5 bg-red-500/20 text-red-300 rounded-lg font-bold block w-fit border border-red-500/30">
                                  ⚠️ Faltó Multa ($3.000)
                                </span>
                              ) : esFueraDeTiempo ? (
                                <span className="px-2 py-0.5 bg-emerald-500/20 text-emerald-300 rounded-lg font-semibold block w-fit">
                                  ✓ Multa Incluida
                                </span>
                              ) : (
                                <span className="text-slate-400">En fecha ordinaria</span>
                              )}
                            </td>
                            <td className="p-3">
                              <a href={t.comprobante_url} target="_blank" rel="noreferrer" className="text-blue-400 underline font-semibold">
                                Ver Imagen
                              </a>
                            </td>
                            <td className="p-3">
                              <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${
                                t.estado === 'aprobado' ? 'bg-emerald-500/20 text-emerald-400' : t.estado === 'rechazado' ? 'bg-red-500/20 text-red-400' : 'bg-amber-500/20 text-amber-400'
                              }`}>
                                {t.estado}
                              </span>
                            </td>
                            <td className="p-3">
                              <div className="flex justify-center gap-1">
                                <button onClick={() => actualizarEstado(t, 'aprobado')} className="p-1.5 bg-emerald-600/30 hover:bg-emerald-600 text-emerald-300 hover:text-white rounded-lg">
                                  <Check className="w-4 h-4" />
                                </button>
                                <button onClick={() => actualizarEstado(t, 'rechazado')} className="p-1.5 bg-red-600/30 hover:bg-red-600 text-red-300 hover:text-white rounded-lg">
                                  <X className="w-4 h-4" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* PRÉSTAMOS: CON LIQUIDACIÓN (MEJORA 4) */}
          {pestanaAdmin === 'prestamos' && (
            <div className="space-y-6">
              <div className="bg-slate-800 p-5 rounded-2xl border border-slate-700">
                <h3 className="font-bold text-white mb-4 flex items-center gap-2">
                  <PlusCircle className="w-5 h-5 text-emerald-400" /> Registrar Nuevo Desembolso de Préstamo
                </h3>
                <form onSubmit={crearPrestamo} className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Nombre del Socio</label>
                    <select
                      value={socioNuevoPrestamo || ''}
                      onChange={(e) => setSocioNuevoPrestamo(Number(e.target.value))}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    >
                      <option value="">Seleccionar socio...</option>
                      {socios.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Valor del Préstamo ($)</label>
                    <input
                      type="text"
                      placeholder="Ej: 500.000"
                      value={montoNuevoPrestamo}
                      onChange={(e) => setMontoNuevoPrestamo(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Interés Mensual (%)</label>
                    <input
                      type="number"
                      step="0.1"
                      placeholder="Ej: 2.0"
                      value={tasaNuevoPrestamo}
                      onChange={(e) => setTasaNuevoPrestamo(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Fecha de Solicitud / Desembolso</label>
                    <input
                      type="date"
                      value={fechaNuevoPrestamo}
                      onChange={(e) => setFechaNuevoPrestamo(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <button type="submit" className="md:col-span-4 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-2.5 rounded-xl text-sm transition mt-1 shadow-md">
                    Guardar Préstamo
                  </button>
                </form>
              </div>

              <div className="bg-slate-800 rounded-2xl border border-slate-700 overflow-x-auto">
                <table className="w-full text-left text-sm text-slate-300">
                  <thead className="bg-slate-700/50 text-slate-400 uppercase text-xs">
                    <tr>
                      <th className="p-3">Socio</th>
                      <th className="p-3">Desembolso</th>
                      <th className="p-3">Fecha Inicio</th>
                      <th className="p-3">Saldo Capital</th>
                      <th className="p-3">Tasa</th>
                      <th className="p-3">Interés / Mes</th>
                      <th className="p-3">Estado</th>
                      <th className="p-3 text-center">Acción</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-700/50">
                    {prestamos.length === 0 ? (
                      <tr><td colSpan={8} className="p-8 text-center text-slate-400">No hay préstamos activos.</td></tr>
                    ) : (
                      prestamos.map((p) => {
                        const intMensual = (Number(p.saldo_capital) * Number(p.tasa_interes_mensual)) / 100;
                        return (
                          <tr key={p.id}>
                            <td className="p-3 font-semibold text-white">{p.socios?.nombre || 'Desconocido'}</td>
                            <td className="p-3">${formatearMoneda(Number(p.monto_prestado))}</td>
                            <td className="p-3">{p.fecha_desembolso}</td>
                            <td className="p-3 font-bold text-amber-400">${formatearMoneda(Number(p.saldo_capital))}</td>
                            <td className="p-3">{p.tasa_interes_mensual}%</td>
                            <td className="p-3 text-emerald-400 font-medium">${formatearMoneda(intMensual)}</td>
                            <td className="p-3">
                              <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${
                                p.estado === 'activo' ? 'bg-amber-500/20 text-amber-400' : 'bg-emerald-500/20 text-emerald-400'
                              }`}>
                                {p.estado}
                              </span>
                            </td>
                            <td className="p-3 text-center">
                              {p.estado === 'activo' && (
                                <button
                                  onClick={() => liquidarPrestamo(p.id)}
                                  className="px-2.5 py-1 bg-slate-700 hover:bg-emerald-600 text-slate-300 hover:text-white rounded-lg text-xs font-medium transition"
                                  title="Liquidar crédito y pasar a Paz y Salvo"
                                >
                                  Paz y Salvo
                                </button>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ACTIVIDADES */}
          {pestanaAdmin === 'actividades' && (
            <div className="space-y-6">
              <div className="bg-slate-800 p-5 rounded-2xl border border-slate-700">
                <h3 className="font-bold text-white mb-4">Crear Nueva Rifa o Actividad Grupal</h3>
                <form onSubmit={crearActividad} className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Nombre de la Actividad</label>
                    <input
                      type="text"
                      placeholder="Ej: Bingo Navideño"
                      value={nombreNuevaActividad}
                      onChange={(e) => setNombreNuevaActividad(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Valor de la Actividad ($)</label>
                    <input
                      type="text"
                      placeholder="Ej: 20.000"
                      value={valorNuevaActividad}
                      onChange={(e) => setValorNuevaActividad(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Costo de Premios ($)</label>
                    <input
                      type="text"
                      placeholder="Ej: 500.000"
                      value={premiosNuevaActividad}
                      onChange={(e) => setPremiosNuevaActividad(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-300 mb-1.5">Fecha Límite</label>
                    <input
                      type="date"
                      value={fechaNuevaActividad}
                      onChange={(e) => setFechaNuevaActividad(e.target.value)}
                      className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white text-sm"
                      required
                    />
                  </div>

                  <button type="submit" className="md:col-span-4 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-2.5 rounded-xl text-sm transition mt-1 shadow-md">
                    Publicar Actividad
                  </button>
                </form>
              </div>

              <div className="space-y-3">
                <h4 className="font-bold text-white text-sm flex items-center gap-2">
                  <CalendarCheck className="w-4 h-4 text-emerald-400" /> Actividades Registradas
                </h4>
                {actividades.length === 0 ? (
                  <div className="p-6 text-center bg-slate-800 rounded-xl border border-slate-700 text-slate-400 text-sm">
                    No hay actividades registradas en el momento. La lista de control aparecerá cuando crees una.
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {actividades.map((act) => {
                      const pagosAct = transaccionesAprobadas.filter((t) => t.desglose_actividad > 0);
                      const totalRecaudado = pagosAct.reduce((acc, curr) => acc + Number(curr.desglose_actividad), 0);
                      const costoPremios = Number(act.costo_premios || 0);
                      const gananciaLimpia = Math.max(0, totalRecaudado - costoPremios);

                      return (
                        <div key={act.id} className="bg-slate-800 p-4 rounded-xl border border-slate-700 space-y-2">
                          <div className="flex justify-between items-start">
                            <div>
                              <h5 className="font-bold text-white text-base">{act.nombre}</h5>
                              <p className="text-xs text-slate-400">Límite: {act.fecha_limite || 'Sin fecha'}</p>
                            </div>
                            <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${
                              act.activa ? 'bg-emerald-500/20 text-emerald-400' : 'bg-slate-700 text-slate-400'
                            }`}>
                              {act.activa ? 'Vigente' : 'Cerrada'}
                            </span>
                          </div>

                          <div className="grid grid-cols-2 gap-2 text-xs border-t border-slate-700/60 pt-2">
                            <div>
                              <span className="text-slate-400 block">Cuota por persona:</span>
                              <span className="font-bold text-white">${formatearMoneda(Number(act.valor_cuota))}</span>
                            </div>
                            <div>
                              <span className="text-slate-400 block">Premios a entregar:</span>
                              <span className="font-bold text-red-400">${formatearMoneda(costoPremios)}</span>
                            </div>
                            <div>
                              <span className="text-slate-400 block">Total Recaudado:</span>
                              <span className="font-bold text-white">${formatearMoneda(totalRecaudado)}</span>
                            </div>
                            <div>
                              <span className="text-emerald-400 block font-semibold">Ganancia Neta Fondo:</span>
                              <span className="font-bold text-emerald-400 text-sm">${formatearMoneda(gananciaLimpia)}</span>
                            </div>
                          </div>

                          <button
                            onClick={() => cambiarEstadoActividad(act.id, act.activa)}
                            className="w-full mt-2 text-xs py-1.5 rounded-lg bg-slate-700 hover:bg-slate-600 text-slate-300 font-medium transition"
                          >
                            {act.activa ? 'Finalizar / Cerrar' : 'Reactivar'}
                          </button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {actividades.length > 0 && (
                <div className="bg-slate-800 rounded-2xl border border-slate-700 p-4">
                  <h4 className="font-bold text-white mb-3 text-sm">
                    Control de Pagos de Actividades ({actividadActiva ? actividadActiva.nombre : 'Última Actividad'})
                  </h4>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
                    {socios.map((s) => {
                      const pagado = transaccionesAprobadas.some(
                        (t) => t.socio_id === s.id && t.desglose_actividad > 0
                      );
                      return (
                        <div key={s.id} className="p-2.5 bg-slate-700/40 rounded-xl border border-slate-600 flex justify-between items-center text-xs">
                          <span className="font-medium text-slate-200">{s.nombre}</span>
                          {pagado ? (
                            <span className="px-2 py-0.5 bg-emerald-500/20 text-emerald-400 rounded-full font-bold">Pagó</span>
                          ) : (
                            <span className="px-2 py-0.5 bg-red-500/20 text-red-400 rounded-full font-bold">Pendiente</span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ATRASADOS */}
          {pestanaAdmin === 'atrasados' && (
            <div className="bg-slate-800 rounded-2xl border border-slate-700 p-5 space-y-5">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-700 pb-4">
                <div>
                  <h3 className="font-bold text-white text-base">Auditoría Quincenal de Ahorro y Multas Pendientes</h3>
                  <p className="text-xs text-slate-400 mt-0.5">Control de Cuota 1 ($50k), Cuota 2 ($50k) y saldo de multas</p>
                </div>

                <div className="flex items-center gap-2">
                  <select
                    value={mesSeleccionado}
                    onChange={(e) => setMesSeleccionado(e.target.value)}
                    className="bg-slate-700 border border-slate-600 rounded-xl px-3 py-2 text-white text-sm font-semibold"
                  >
                    {MESES.map((m) => <option key={m.valor} value={m.valor}>{m.nombre}</option>)}
                  </select>
                  <select
                    value={anioSeleccionado}
                    onChange={(e) => setAnioSeleccionado(e.target.value)}
                    className="bg-slate-700 border border-slate-600 rounded-xl px-3 py-2 text-white text-sm font-semibold"
                  >
                    <option value="2026">2026</option>
                    <option value="2027">2027</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {sociosDetalleMes.map((s) => {
                  const tieneMultasPendientes = s.multasPendientesValor > 0;
                  const requiereCobro = s.estado !== 'al_dia' || tieneMultasPendientes;

                  let mensajeCobro = `Hola ${s.nombre}, te escribimos del Fondo de Ahorro para comentarte tu estado de cuenta: `;
                  if (s.estado !== 'al_dia') {
                    mensajeCobro += `Tienes pendiente el pago de tu ${s.textoAtraso} de ${nombreMes} ${anioSeleccionado}. `;
                  }
                  if (tieneMultasPendientes) {
                    mensajeCobro += `Adicionalmente, registras $${formatearMoneda(s.multasPendientesValor)} COP pendientes de saldo en multas.`;
                  }

                  return (
                    <div key={s.id} className="p-3.5 bg-slate-700/40 border border-slate-600/70 rounded-xl flex flex-col justify-between gap-3 text-sm">
                      <div className="flex justify-between items-start">
                        <div>
                          <p className="font-semibold text-white">{s.nombre}</p>
                          <div className="flex gap-2 text-xs mt-1">
                            <span className={s.pagoCuota1 ? 'text-emerald-400 font-medium' : 'text-red-400'}>
                              Cuota 1: {s.pagoCuota1 ? '✓ Pagada' : '✗ Pendiente'}
                            </span>
                            <span className="text-slate-500">•</span>
                            <span className={s.pagoCuota2 ? 'text-emerald-400 font-medium' : 'text-red-400'}>
                              Cuota 2: {s.pagoCuota2 ? '✓ Pagada' : '✗ Pendiente'}
                            </span>
                          </div>
                        </div>

                        {tieneMultasPendientes ? (
                          <span className="px-2 py-1 bg-red-500/20 text-red-300 border border-red-500/30 rounded-lg text-xs font-bold text-right">
                            Debe ${formatearMoneda(s.multasPendientesValor)} en Multas
                          </span>
                        ) : s.totalMultasAbonadas > 0 ? (
                          <span className="px-2 py-1 bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 rounded-lg text-xs font-medium">
                            ✓ Multas Saldadas
                          </span>
                        ) : null}
                      </div>

                      <div className="flex justify-between items-center pt-2 border-t border-slate-700">
                        <span className="text-xs text-slate-400">
                          {s.estado === 'al_dia' && !tieneMultasPendientes ? 'Al día en cuotas y multas' : 'Saldo o cuota pendiente'}
                        </span>

                        {requiereCobro ? (
                          <a
                            href={`https://wa.me/?text=${encodeURIComponent(mensajeCobro)}`}
                            target="_blank"
                            rel="noreferrer"
                            className="px-3 py-1.5 bg-emerald-600/20 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-emerald-500/30 text-xs rounded-lg font-medium transition"
                          >
                            Cobrar WhatsApp
                          </a>
                        ) : (
                          <span className="px-2 py-1 bg-emerald-500/20 text-emerald-400 rounded-lg text-xs font-bold">
                            Al día
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      ) : modoPublico === 'consultar' ? (
        /* MEJORA 1: PORTAL DE CONSULTA PÚBLICA DEL SOCIO (CERO LLAMADAS) */
        <div className="max-w-xl mx-auto bg-slate-800 rounded-2xl shadow-xl p-6 md:p-8 border border-slate-700 space-y-6">
          <div className="text-center">
            <h2 className="text-2xl font-bold text-white flex items-center justify-center gap-2">
              <Search className="w-6 h-6 text-blue-400" /> Consulta de Estado de Cuenta
            </h2>
            <p className="text-xs text-slate-400 mt-1">Revisa tus ahorros, comprobantes y deudas al instante</p>
          </div>

          <div>
            <label className="block text-xs font-semibold text-slate-300 mb-1.5">Selecciona tu Nombre</label>
            <select
              value={socioConsultaId || ''}
              onChange={(e) => setSocioConsultaId(Number(e.target.value))}
              className="w-full bg-slate-700 border border-slate-600 rounded-xl p-3 text-white text-sm font-semibold focus:outline-none focus:border-blue-500"
            >
              {socios.map((s) => (
                <option key={s.id} value={s.id}>{s.nombre}</option>
              ))}
            </select>
          </div>

          {socioConsultaObj && (
            <div className="space-y-4">
              {/* Tarjetas Resumen */}
              <div className="grid grid-cols-2 gap-3">
                <div className="bg-emerald-500/10 border border-emerald-500/30 p-4 rounded-xl">
                  <span className="text-xs text-emerald-300 block mb-1">Total Ahorrado 2027</span>
                  <span className="text-xl font-extrabold text-emerald-400">${formatearMoneda(totalAhorradoConsulta)}</span>
                </div>
                <div className="bg-slate-700/40 border border-slate-600 p-4 rounded-xl">
                  <span className="text-xs text-slate-400 block mb-1">Saldo Multas</span>
                  <span className={`text-xl font-extrabold ${saldoMultasConsulta > 0 ? 'text-red-400' : 'text-white'}`}>
                    ${formatearMoneda(saldoMultasConsulta)}
                  </span>
                </div>
              </div>

              {/* Estado Préstamo si tiene */}
              {prestamoConsulta ? (
                <div className="bg-amber-500/10 border border-amber-500/30 p-4 rounded-xl space-y-1 text-xs">
                  <div className="flex justify-between items-center text-amber-300 font-bold mb-1">
                    <span>Préstamo Activo Vigente</span>
                    <span>{prestamoConsulta.tasa_interes_mensual}% mensual</span>
                  </div>
                  <div className="flex justify-between text-slate-300">
                    <span>Deuda Capital Pendiente:</span>
                    <span className="font-bold text-white">${formatearMoneda(Number(prestamoConsulta.saldo_capital))}</span>
                  </div>
                </div>
              ) : (
                <div className="p-3 bg-slate-700/20 border border-slate-700 rounded-xl text-center text-xs text-slate-400">
                  Sin créditos o préstamos activos al momento.
                </div>
              )}

              {/* Historial de Comprobantes Recientes */}
              <div className="space-y-2 pt-2">
                <h4 className="text-xs font-bold text-slate-300 uppercase tracking-wider">Últimos Reportes Enviados</h4>
                <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                  {pagosConsulta.length === 0 ? (
                    <p className="text-xs text-slate-500 text-center py-4">No has registrado pagos aún.</p>
                  ) : (
                    pagosConsulta.map((p) => (
                      <div key={p.id} className="p-2.5 bg-slate-700/40 rounded-lg border border-slate-600/70 flex justify-between items-center text-xs">
                        <div>
                          <p className="font-semibold text-white">{p.concepto}</p>
                          <p className="text-[10px] text-slate-400">{new Date(p.fecha_reporte).toLocaleDateString()}</p>
                        </div>
                        <div className="text-right">
                          <span className="font-bold text-emerald-400 block">${formatearMoneda(Number(p.monto_reportado))}</span>
                          <span className={`text-[10px] font-semibold ${p.estado === 'aprobado' ? 'text-emerald-400' : 'text-amber-400'}`}>
                            {p.estado === 'aprobado' ? '✓ Aprobado' : '⏳ En revisión'}
                          </span>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      ) : (
        /* VISTA REPORTE DE PAGO Bre-B */
        <div className="max-w-xl mx-auto bg-slate-800 rounded-2xl shadow-xl p-6 md:p-8 border border-slate-700">
          <div className="text-center mb-6">
            <h2 className="text-2xl font-bold text-white">Reporte de Pago Bre-B</h2>
            <p className="text-sm text-slate-400 mt-1">Fondo de Ahorro</p>
          </div>

          {resumenExito ? (
            <div className="bg-slate-900 border border-emerald-500/40 rounded-2xl p-6 space-y-4">
              <div className="flex items-center gap-3 text-emerald-400">
                <CheckCircle2 className="w-8 h-8" />
                <div>
                  <h3 className="font-bold text-lg text-white">¡Comprobante Enviado!</h3>
                  <p className="text-xs text-slate-400">En revisión para conciliación en auditoría.</p>
                </div>
              </div>

              <div className="bg-slate-800 p-4 rounded-xl space-y-2 text-sm border border-slate-700">
                <div className="flex justify-between"><span className="text-slate-400">Socio:</span><span className="font-semibold text-white">{resumenExito.nombre}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">Conceptos:</span><span className="font-semibold text-white text-right">{resumenExito.conceptos}</span></div>
                <div className="flex justify-between"><span className="text-slate-400">Fecha:</span><span className="font-semibold text-white">{resumenExito.fecha}</span></div>
                <div className="flex justify-between border-t border-slate-700 pt-2"><span className="text-slate-400">Total Consignado:</span><span className="font-bold text-emerald-400 text-base">${formatearMoneda(resumenExito.total)} COP</span></div>
              </div>

              {resumenExito.conMulta && (
                <div className="p-3 bg-amber-500/10 border border-amber-500/30 rounded-xl text-amber-300 text-xs">
                  Recuerda: Consignaste en fecha con recargo. Si no incluiste la multa pactada ($3.000), se cobrará en el siguiente corte.
                </div>
              )}

              <button onClick={() => setResumenExito(null)} className="w-full bg-slate-700 py-2.5 rounded-xl text-sm font-semibold">Reportar otro pago</button>
            </div>
          ) : (
            <form onSubmit={enviarPago} className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-300 mb-1">Tu Nombre</label>
                <select
                  value={socioSeleccionado || ''}
                  onChange={(e) => setSocioSeleccionado(Number(e.target.value))}
                  className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white focus:outline-none focus:border-emerald-500"
                  required
                >
                  <option value="">Selecciona tu nombre...</option>
                  {socios.map((s) => <option key={s.id} value={s.id}>{s.nombre}</option>)}
                </select>
              </div>

              {prestamoActivoSocio && (
                <div className="p-4 bg-blue-500/10 border border-blue-500/30 rounded-xl space-y-1.5 text-xs text-blue-200">
                  <p className="font-bold text-blue-300">
                    Información de tu Préstamo Activo
                  </p>
                  <div className="flex justify-between"><span>Deuda de Capital:</span><span className="font-semibold text-white">${formatearMoneda(Number(prestamoActivoSocio.saldo_capital))}</span></div>
                  <div className="flex justify-between"><span>Tasa pactada:</span><span>{prestamoActivoSocio.tasa_interes_mensual}% mensual</span></div>
                  <div className="flex justify-between"><span>Meses transcurridos:</span><span>{mesesInteresCalculados} mes(es)</span></div>
                  <div className="flex justify-between border-t border-blue-500/30 pt-1 font-bold text-amber-300"><span>Intereses causados al corte:</span><span>${formatearMoneda(totalInteresEstimado)}</span></div>
                </div>
              )}

              <div>
                <label className="block text-sm font-medium text-slate-300 mb-1">Fecha de Consignación</label>
                <input
                  type="date"
                  value={fechaPago}
                  onChange={(e) => setFechaPago(e.target.value)}
                  className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2.5 text-white focus:outline-none focus:border-emerald-500"
                  required
                />
              </div>

              {/* CUOTA DE AHORRO */}
              <div className="bg-slate-700/40 p-4 rounded-xl border border-slate-600 space-y-3">
                <div className="flex justify-between items-center">
                  <label className="text-sm font-bold text-white flex items-center gap-2">
                    <input
                      type="checkbox"
                      checked={pagarAhorro}
                      onChange={(e) => setPagarAhorro(e.target.checked)}
                      className="w-4 h-4 rounded text-emerald-600 focus:ring-0"
                    />
                    Ahorro Mensual ($100.000)
                  </label>
                  <span className="text-xs font-bold text-emerald-400">
                    Subtotal: ${formatearMoneda(montoAhorroCalculado)}
                  </span>
                </div>

                {pagarAhorro && (
                  <div className="grid grid-cols-2 gap-2 pt-1 pl-6">
                    <button
                      type="button"
                      onClick={() => setCuota1Seleccionada(!cuota1Seleccionada)}
                      className={`p-2 rounded-lg border text-xs font-medium flex items-center justify-between ${
                        cuota1Seleccionada ? 'bg-emerald-500/20 border-emerald-500 text-white' : 'bg-slate-800 border-slate-600 text-slate-400'
                      }`}
                    >
                      <span>Cuota 1 (Quincena 1)</span>
                      <span className="font-bold">$50.000</span>
                    </button>

                    <button
                      type="button"
                      onClick={() => setCuota2Seleccionada(!cuota2Seleccionada)}
                      className={`p-2 rounded-lg border text-xs font-medium flex items-center justify-between ${
                        cuota2Seleccionada ? 'bg-emerald-500/20 border-emerald-500 text-white' : 'bg-slate-800 border-slate-600 text-slate-400'
                      }`}
                    >
                      <span>Cuota 2 (Quincena 2)</span>
                      <span className="font-bold">$50.000</span>
                    </button>
                  </div>
                )}
              </div>

              {/* OTROS CONCEPTOS */}
              <div className="space-y-2">
                <label className="block text-xs font-semibold text-slate-400 uppercase tracking-wider">
                  Otros Conceptos (Opcional)
                </label>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => setPagarPrestamo(!pagarPrestamo)}
                    className={`p-2 rounded-xl border text-xs font-medium text-center ${
                      pagarPrestamo ? 'bg-blue-500/20 border-blue-500 text-white' : 'bg-slate-700/40 border-slate-600 text-slate-300'
                    }`}
                  >
                    Préstamo
                  </button>
                  <button
                    type="button"
                    onClick={() => setPagarMulta(!pagarMulta)}
                    className={`p-2 rounded-xl border text-xs font-medium text-center ${
                      pagarMulta ? 'bg-red-500/20 border-red-500 text-white' : 'bg-slate-700/40 border-slate-600 text-slate-300'
                    }`}
                  >
                    Multa ($3.000)
                  </button>
                  <button
                    type="button"
                    onClick={() => setPagarActividad(!pagarActividad)}
                    className={`p-2 rounded-xl border text-xs font-medium text-center ${
                      pagarActividad ? 'bg-purple-500/20 border-purple-500 text-white' : 'bg-slate-700/40 border-slate-600 text-slate-300'
                    }`}
                  >
                    Actividad
                  </button>
                </div>
              </div>

              {(pagarPrestamo || pagarMulta || pagarActividad) && (
                <div className="space-y-3 bg-slate-700/20 p-4 rounded-xl border border-slate-700">
                  {pagarPrestamo && (
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="block text-xs text-amber-300 mb-1">Pago Interés ($)</label>
                        <input
                          type="text"
                          value={montoInteres}
                          onChange={(e) => setMontoInteres(e.target.value)}
                          className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs font-semibold"
                          placeholder="0"
                        />
                      </div>
                      <div>
                        <label className="block text-xs text-blue-300 mb-1">Abono Capital ($)</label>
                        <input
                          type="text"
                          value={montoCapital}
                          onChange={(e) => setMontoCapital(e.target.value)}
                          className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs font-semibold"
                          placeholder="0"
                        />
                      </div>
                    </div>
                  )}

                  {pagarMulta && (
                    <div>
                      <label className="block text-xs text-red-300 mb-1">Monto de Multa ($)</label>
                      <input
                        type="text"
                        value={montoMulta}
                        onChange={(e) => setMontoMulta(e.target.value)}
                        className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs font-semibold"
                        placeholder="3.000"
                      />
                    </div>
                  )}

                  {pagarActividad && (
                    <div>
                      <label className="block text-xs text-purple-300 mb-1">
                        Aporte Actividad {actividadActiva ? `("${actividadActiva.nombre}")` : ''} ($)
                      </label>
                      <input
                        type="text"
                        value={montoActividad}
                        onChange={(e) => setMontoActividad(e.target.value)}
                        className="w-full bg-slate-700 border border-slate-600 rounded-lg p-2 text-white text-xs font-semibold"
                        placeholder={actividadActiva ? formatearMoneda(actividadActiva.valor_cuota) : '20.000'}
                      />
                    </div>
                  )}
                </div>
              )}

              <div className="border-t border-slate-600 pt-3 flex justify-between items-center">
                <span className="text-sm font-bold text-white">Total a consignar por Bre-B:</span>
                <span className="text-lg font-extrabold text-emerald-400">
                  ${formatearMoneda(totalAConsignar)} COP
                </span>
              </div>

              <div>
                <label className="block text-sm font-medium text-slate-300 mb-1">Adjuntar Comprobante de pago</label>
                <div className="border-2 border-dashed border-slate-600 rounded-lg p-4 bg-slate-700/30 text-center cursor-pointer">
                  <label className="cursor-pointer block">
                    <Upload className="w-6 h-6 text-slate-400 mx-auto mb-1" />
                    <span className="text-xs text-slate-300 block">
                      {archivo ? archivo.name : 'Toca para subir captura o foto'}
                    </span>
                    <input type="file" accept="image/*" onChange={(e) => setArchivo(e.target.files?.[0] || null)} className="hidden" />
                  </label>
                </div>
              </div>

              {errorForm && <p className="text-xs text-red-400 bg-red-500/10 p-2.5 rounded-lg border border-red-500/30">{errorForm}</p>}

              <button
                type="submit"
                disabled={cargandoForm}
                className="w-full bg-emerald-600 hover:bg-emerald-500 font-semibold py-3 rounded-xl transition text-white shadow-lg disabled:opacity-50"
              >
                {cargandoForm ? 'Enviando...' : 'Enviar Reporte'}
              </button>
            </form>
          )}
        </div>
      )}
    </main>
  );
}