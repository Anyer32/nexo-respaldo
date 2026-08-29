// generar-respaldo.js
// Genera respaldo-nexo.xlsx con 2 hojas: Catálogo (código, descripción, categoría, cantidad)
// y GSI (rotación de ventas del período). Pensado para correr en GitHub Actions.

const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;
const RESPALDO_USUARIO = process.env.RESPALDO_USUARIO;
const RESPALDO_PASSWORD = process.env.RESPALDO_PASSWORD;

// Rango de fechas para la hoja GSI: del 1 del mes actual a hoy (igual que en la app).
// Si más adelante querés otra ventana (ej. "últimos 30 días"), cambiá esta función.
function rangoGsi() {
  const hoy = new Date();
  const desde = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().slice(0, 10);
  const hasta = hoy.toISOString().slice(0, 10);
  return { desde, hasta };
}

// Trae TODAS las filas de una tabla, en páginas de 1000 (el máximo que entrega Supabase por consulta).
async function obtenerTodo(supabase, tabla, columnas, ordenarPor) {
  const TAM_PAGINA = 1000;
  let desde = 0;
  let todas = [];
  while (true) {
    const { data, error } = await supabase
      .from(tabla)
      .select(columnas)
      .order(ordenarPor, { ascending: true })
      .range(desde, desde + TAM_PAGINA - 1);
    if (error) throw new Error(`Leyendo ${tabla}: ${error.message}`);
    todas = todas.concat(data);
    if (!data.length || data.length < TAM_PAGINA) break;
    desde += TAM_PAGINA;
  }
  return todas;
}

async function main() {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY || !RESPALDO_USUARIO || !RESPALDO_PASSWORD) {
    throw new Error('Faltan variables de entorno: SUPABASE_URL, SUPABASE_ANON_KEY, RESPALDO_USUARIO, RESPALDO_PASSWORD.');
  }

  const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  const email = RESPALDO_USUARIO.includes('@') ? RESPALDO_USUARIO : `${RESPALDO_USUARIO}@tualiado.local`;
  const { error: errorLogin } = await supabase.auth.signInWithPassword({ email, password: RESPALDO_PASSWORD });
  if (errorLogin) throw new Error(`No se pudo iniciar sesión con la cuenta de respaldo: ${errorLogin.message}`);

  console.log('Sesión iniciada. Leyendo tiendas...');
  const tiendas = await obtenerTodo(supabase, 'tiendas', 'id, nombre', 'nombre');
  if (!tiendas.length) throw new Error('No se encontró ninguna tienda — revisá los permisos de la cuenta de respaldo.');

  console.log(`Leyendo catálogo de ${tiendas.length} tienda(s)...`);
  const productos = await obtenerTodo(
    supabase, 'productos', 'codigo, descripcion, stock, tienda_id, categorias(nombre)', 'codigo'
  );

  const { desde, hasta } = rangoGsi();
  console.log(`Calculando GSI del ${desde} al ${hasta}...`);
  const filasGsi = [];
  for (const tienda of tiendas) {
    const { data, error } = await supabase.rpc('gsi_rotacion_periodo', {
      p_desde: desde, p_hasta: hasta, p_tienda_id: tienda.id
    });
    if (error) throw new Error(`GSI de "${tienda.nombre}": ${error.message}`);
    (data || []).forEach((f) => filasGsi.push({ ...f, tienda_nombre: tienda.nombre }));
  }

  console.log('Armando el Excel...');
  const nombrePorTienda = new Map(tiendas.map((t) => [t.id, t.nombre]));
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Nexo';
  workbook.created = new Date();

  const hojaCatalogo = workbook.addWorksheet('Catálogo');
  hojaCatalogo.columns = [
    { header: 'Tienda', key: 'tienda', width: 22 },
    { header: 'Código', key: 'codigo', width: 16 },
    { header: 'Descripción', key: 'descripcion', width: 45 },
    { header: 'Categoría', key: 'categoria', width: 22 },
    { header: 'Cantidad', key: 'cantidad', width: 12 }
  ];
  hojaCatalogo.getRow(1).font = { bold: true };
  hojaCatalogo.views = [{ state: 'frozen', ySplit: 1 }];
  productos.forEach((p) => {
    hojaCatalogo.addRow({
      tienda: nombrePorTienda.get(p.tienda_id) || '',
      codigo: p.codigo,
      descripcion: p.descripcion,
      categoria: p.categorias?.nombre || '',
      cantidad: p.stock
    });
  });
  hojaCatalogo.autoFilter = { from: 'A1', to: `E${productos.length + 1}` };

  const hojaGsi = workbook.addWorksheet('GSI');
  hojaGsi.columns = [
    { header: 'Tienda', key: 'tienda', width: 22 },
    { header: 'Código', key: 'codigo', width: 16 },
    { header: 'Descripción', key: 'descripcion', width: 45 },
    { header: 'Categoría', key: 'categoria', width: 22 },
    { header: `Ventas (${desde} a ${hasta})`, key: 'ventas', width: 18 },
    { header: 'Stock', key: 'stock', width: 10 },
    { header: 'Participación %', key: 'participacion', width: 16 },
    { header: 'Sin rotación', key: 'sinRotacion', width: 14 }
  ];
  hojaGsi.getRow(1).font = { bold: true };
  hojaGsi.views = [{ state: 'frozen', ySplit: 1 }];
  filasGsi.forEach((f) => {
    hojaGsi.addRow({
      tienda: f.tienda_nombre,
      codigo: f.codigo,
      descripcion: f.descripcion,
      categoria: f.categoria_nombre || '',
      ventas: f.ventas,
      stock: f.stock,
      participacion: f.participacion,
      sinRotacion: f.sin_rotacion ? 'Sí' : ''
    });
  });
  hojaGsi.autoFilter = { from: 'A1', to: `H${filasGsi.length + 1}` };

  await workbook.xlsx.writeFile('respaldo-nexo.xlsx');
  console.log(`Listo: respaldo-nexo.xlsx (${productos.length} productos, ${filasGsi.length} filas de GSI).`);
}

main().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
