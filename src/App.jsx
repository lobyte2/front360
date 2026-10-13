import React, { useState, useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import { ShoppingCart, User, Wine, SignOut, ArrowRight, Trash, Package, CheckCircle, WarningCircle, Info, X, ClockCounterClockwise, ShieldCheck, ArrowClockwise, EnvelopeSimple, Truck, Database, Lightning, CaretLeft, CaretRight, Broom } from '@phosphor-icons/react';

const API_GATEWAY_URL = "https://ap96rduot1.execute-api.us-east-1.amazonaws.com/api";
const USUARIO_ID = 1; 

function App() {
    const { instance, accounts } = useMsal();
    const isAuthenticated = useIsAuthenticated();
    
    const [vista, setVista] = useState('catalogo');
    const [productos, setProductos] = useState([]);
    const [carrito, setCarrito] = useState(null);
    const [perfil, setPerfil] = useState(null);
    const [notificaciones, setNotificaciones] = useState([]);
    const [filtroAuditoria, setFiltroAuditoria] = useState('TODOS');
    const [paginaAuditoria, setPaginaAuditoria] = useState(1);
    const [logsPorPagina, setLogsPorPagina] = useState(10);
    const [menuLimpiarAbierto, setMenuLimpiarAbierto] = useState(false);
    const [auditoriaLogs, setAuditoriaLogs] = useState([
        {
            id: 'init-1',
            timestamp: new Date().toLocaleTimeString(),
            origen: 'SISTEMA',
            accion: 'INICIALIZACION',
            descripcion: 'Ecosistema de microservicios Pedidos360 conectado a RabbitMQ',
            nivel: 'INFO',
            detalles: { broker: '34.236.39.87:5672', eureka: '52.205.162.211:8761' }
        },
        {
            id: 'init-2',
            timestamp: new Date().toLocaleTimeString(),
            origen: 'MS-PRODUCTOS',
            accion: 'CATALOGO_LISTO',
            descripcion: 'Inventario de 5 productos cargado en memoria y disponible para compras',
            nivel: 'INFO',
            detalles: { productos_activos: 5 }
        }
    ]);

    const mostrarNotificacion = (mensaje, tipo = 'exito') => {
        const id = Date.now() + '-' + Math.random().toString(36).substring(2, 6);
        setNotificaciones(prev => [...prev, { id, mensaje, tipo }]);

        setTimeout(() => {
            setNotificaciones(prev => prev.filter(n => n.id !== id));
        }, 5000); // 5 segundos para que haya buen tiempo de lectura
    };

    const cerrarNotificacion = (id) => {
        setNotificaciones(prev => prev.filter(n => n.id !== id));
    };

    const registrarAuditoria = async (origen, accion, descripcion, nivel = 'INFO', detalles = {}) => {
        const nuevoLog = {
            id: Date.now() + '-' + Math.random().toString(36).substring(2, 6),
            timestamp: new Date().toLocaleTimeString(),
            origen,
            accion,
            descripcion,
            nivel,
            detalles
        };
        setAuditoriaLogs(prev => [nuevoLog, ...prev]);

        try {
            await peticionAutenticada('/auditoria', {
                method: 'POST',
                body: JSON.stringify(nuevoLog)
            });
        } catch (e) {
            // Silencioso si el API Gateway aun no tiene mapeada la ruta
        }
    };

    const formatearTimestamp = (ts) => {
        if (!ts) return '';
        if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(ts)) return ts;
        try {
            // Reemplazar espacio por T para parsear correctamente fecha ISO/SQL
            const fechaLimpia = ts.includes('T') ? ts : ts.replace(' ', 'T');
            const d = new Date(fechaLimpia);
            if (!isNaN(d.getTime())) {
                return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            }
        } catch (e) {
            // Ignorar y retornar tal cual
        }
        return ts;
    };

    const cambiarFiltroAuditoria = (tag) => {
        setFiltroAuditoria(tag);
        setPaginaAuditoria(1);
    };

    const limpiarAuditoria = async (modo = '10') => {
        setMenuLimpiarAbierto(false);
        let nuevaLista = [...auditoriaLogs];
        let cantidad = 10;
        let texto = "";

        if (modo === '10') {
            cantidad = Math.min(10, nuevaLista.length);
            nuevaLista = nuevaLista.slice(cantidad);
            texto = `Se eliminaron los ${cantidad} registros más recientes.`;
        } else if (modo === 'pagina') {
            const logsFiltrados = auditoriaLogs.filter(item => 
                filtroAuditoria === 'TODOS' || (item.origen && item.origen.toUpperCase() === filtroAuditoria)
            );
            const indexInicio = (paginaAuditoria - 1) * logsPorPagina;
            const logsPagina = logsFiltrados.slice(indexInicio, indexInicio + logsPorPagina);
            const idsPagina = new Set(logsPagina.map(l => l.id));
            cantidad = idsPagina.size;
            nuevaLista = nuevaLista.filter(l => !idsPagina.has(l.id));
            texto = `Se eliminaron ${cantidad} registros de la página actual.`;
        } else if (modo === 'todos') {
            cantidad = nuevaLista.length;
            nuevaLista = [];
            texto = "Se eliminaron todos los registros de auditoría.";
        }

        setAuditoriaLogs(nuevaLista);
        setPaginaAuditoria(1);
        mostrarNotificacion(texto, 'info');

        try {
            const query = modo === 'todos' ? '' : `?cantidad=${cantidad}`;
            await peticionAutenticada(`/auditoria${query}`, { method: 'DELETE' });
        } catch (err) {
            // Silencioso si API Gateway no implementa DELETE
        }
    };

    const obtenerAuditoria = async () => {
        try {
            const data = await peticionAutenticada('/auditoria');
            if (Array.isArray(data) && data.length > 0) {
                setAuditoriaLogs(data);
            }
        } catch (error) {
            // Mantiene los logs acumulados localmente
        }
        setPaginaAuditoria(1);
        setVista('auditoria');
    };

    useEffect(() => {
        if (isAuthenticated) {
            obtenerProductos();
        }
    }, [isAuthenticated]);

    const obtenerToken = async () => {
        const response = await instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        });
        return response.accessToken;
    };

    const peticionAutenticada = async (endpoint, options = {}) => {
        const token = await obtenerToken();
        const res = await fetch(`${API_GATEWAY_URL}${endpoint}`, {
            ...options,
            headers: {
                ...options.headers,
                "Authorization": `Bearer ${token}`,
                "Content-Type": "application/json"
            }
        });
        if (!res.ok) throw new Error(`Error HTTP: ${res.status}`);
        const text = await res.text();
        return text ? JSON.parse(text) : null;
    };

    const obtenerProductos = async () => {
        try {
            const data = await peticionAutenticada('/productos');
            setProductos(data);
            setVista('catalogo');
        } catch (error) {
            console.error("Error al cargar el catálogo:", error);
        }
    };

    const obtenerCarrito = async () => {
        try {
            const data = await peticionAutenticada(`/carrito/${USUARIO_ID}`);
            setCarrito(data);
            setVista('carrito');
        } catch (error) {
            console.error("Error al cargar carrito:", error);
        }
    };

    const agregarAlCarrito = async (articulo) => {
        const id = articulo.identificador || articulo.id;
        const prod = productos.find(p => (p.identificador || p.id) === id);
        const stockDisponible = prod ? (prod.inventario ?? 0) : 0;

        const cantidadEnCarrito = carrito?.articulos
            ?.filter(item => item.idProducto === id)
            ?.reduce((sum, item) => sum + item.cantidad, 0) || 0;

        if (stockDisponible <= 0) {
            mostrarNotificacion(`"${articulo.nombre || 'Producto'}" está agotado en bodega.`, 'error');
            return;
        }

        if (cantidadEnCarrito + 1 > stockDisponible) {
            mostrarNotificacion(`Tope de stock alcanzado. Solo quedan ${stockDisponible} unidad(es) de "${articulo.nombre}".`, 'error');
            return;
        }

        try {
            await peticionAutenticada(`/carrito/${USUARIO_ID}/agregar`, {
                method: 'POST',
                body: JSON.stringify({
                    idProducto: id, 
                    cantidad: 1 
                })
            });
            mostrarNotificacion(`"${articulo.nombre || 'Producto'}" agregado al carrito 🍷`, 'exito');
            registrarAuditoria('FRONTEND', 'PRODUCTO_AGREGADO', `Se agregó al carrito: "${articulo.nombre || 'Producto'}"`, 'SUCCESS', { productoId: id, enCarrito: cantidadEnCarrito + 1, stockTotal: stockDisponible });
            obtenerCarrito();
        } catch (error) {
            console.error("Error al agregar producto:", error);
            mostrarNotificacion("Error al agregar el producto al carrito.", "error");
            registrarAuditoria('FRONTEND', 'ERROR_CARRITO', `Fallo al agregar producto "${articulo.nombre}"`, 'ERROR');
        }
    };

    const limpiarCarrito = async (mostrarToast = true) => {
        try {
            await peticionAutenticada(`/carrito/${USUARIO_ID}/limpiar`, { method: 'DELETE' });
            if (mostrarToast) {
                mostrarNotificacion("Carrito limpiado correctamente.", "info");
                registrarAuditoria('FRONTEND', 'CARRITO_LIMPIADO', "El usuario vació manualmente el carrito de compras", 'WARN');
            }
            obtenerCarrito();
        } catch (error) {
            console.error("Error al limpiar carrito:", error);
            if (mostrarToast) {
                mostrarNotificacion("Error al limpiar el carrito.", "error");
            }
        }
    };

    const obtenerPerfil = async () => {
        try {
            const username = accounts[0]?.username;
            const data = await peticionAutenticada(`/usuarios/perfil/${username}`);
            setPerfil(data);
            setVista('perfil');
        } catch (error) {
            console.error("Error al cargar perfil:", error);
        }
    };

    const procederAlPago = async () => {
        if (!carrito || !carrito.articulos || carrito.articulos.length === 0) return;
        
        // Validar si algún producto supera el stock disponible en bodega
        const articulosSinStock = carrito.articulos.filter(item => {
            const prod = productos.find(p => (p.identificador || p.id) === item.idProducto);
            const stock = prod ? (prod.inventario ?? 0) : 0;
            return item.cantidad > stock || stock <= 0;
        });

        if (articulosSinStock.length > 0) {
            mostrarNotificacion("No puedes finalizar la compra: algunos productos en tu orden no tienen stock suficiente en bodega.", "error");
            return;
        }

        const total = carrito.articulos.reduce((sum, item) => {
            const prod = productos.find(p => (p.identificador || p.id) === item.idProducto);
            return sum + (prod ? prod.precio * item.cantidad : 0);
        }, 0);

        const itemsOrden = carrito.articulos.map(item => ({
            productoId: item.idProducto,
            cantidad: item.cantidad
        }));

        try {
            await peticionAutenticada('/ordenes/nueva', {
                method: 'POST',
                body: JSON.stringify({
                    usuarioId: accounts[0]?.username || "anonimo",
                    items: itemsOrden,
                    total: total
                })
            });
            mostrarNotificacion("¡Orden creada con éxito! RabbitMQ procesando notificaciones, despacho y stock...", "exito");
            
            await limpiarCarrito(false);

            // Recargar productos después de que RabbitMQ y Supabase descuenten el stock
            setTimeout(() => {
                obtenerProductos();
            }, 1000);
        } catch (error) {
            console.error("Error al crear la orden:", error);
            mostrarNotificacion("Hubo un error al procesar el pago.", "error");
            registrarAuditoria('MS-ORDENES', 'ERROR_PAGO', 'Error al procesar la orden de compra', 'ERROR');
        }
    };

    const handleLogin = () => instance.loginRedirect(loginRequest);
    const handleLogout = () => instance.logoutRedirect();

    if (!isAuthenticated) {
        return (
            <div className="min-h-[100dvh] bg-zinc-950 text-zinc-50 flex flex-col items-center justify-center p-6 font-sans">
                <div className="w-full max-w-md flex flex-col items-center text-center space-y-8">
                    <div className="p-4 bg-zinc-900 rounded-2xl">
                        <Wine size={48} weight="duotone" className="text-emerald-400" />
                    </div>
                    <div className="space-y-3">
                        <h1 className="text-4xl md:text-5xl font-medium tracking-tight">Pedidos360</h1>
                        <p className="text-zinc-400 text-lg">Acceso exclusivo al catálogo de bodega.</p>
                    </div>
                    <button 
                        onClick={handleLogin}
                        className="w-full flex items-center justify-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-zinc-950 font-medium px-6 py-4 rounded-xl transition-all active:scale-[0.98]"
                    >
                        Ingresar con Azure AD <ArrowRight size={20} />
                    </button>
                </div>
            </div>
        );
    }

    return (
        <div className="min-h-[100dvh] bg-zinc-950 text-zinc-50 font-sans selection:bg-emerald-500/30 selection:text-emerald-200">
            {/* Header */}
            <header className="sticky top-0 z-50 w-full border-b border-zinc-800/50 bg-zinc-950/80 backdrop-blur-md">
                <div className="max-w-7xl mx-auto px-6 h-20 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-900 rounded-lg">
                            <Wine size={24} weight="duotone" className="text-emerald-400" />
                        </div>
                        <span className="text-xl font-medium tracking-tight">Pedidos360</span>
                    </div>
                    
                    <nav className="hidden md:flex items-center gap-8">
                        <button onClick={obtenerProductos} className={`text-sm font-medium transition-colors ${vista === 'catalogo' ? 'text-emerald-400' : 'text-zinc-400 hover:text-zinc-100'}`}>Catálogo</button>
                        <button onClick={obtenerCarrito} className={`text-sm font-medium transition-colors ${vista === 'carrito' ? 'text-emerald-400' : 'text-zinc-400 hover:text-zinc-100'}`}>Carrito</button>
                        <button onClick={obtenerPerfil} className={`text-sm font-medium transition-colors ${vista === 'perfil' ? 'text-emerald-400' : 'text-zinc-400 hover:text-zinc-100'}`}>Perfil</button>
                        <button onClick={obtenerAuditoria} className={`text-sm font-medium transition-colors ${vista === 'auditoria' ? 'text-emerald-400' : 'text-zinc-400 hover:text-zinc-100'}`}>Auditoría</button>
                    </nav>

                    <div className="flex items-center gap-4">
                        <div className="md:hidden flex gap-4">
                            <button onClick={obtenerProductos} className={`${vista === 'catalogo' ? 'text-emerald-400' : 'text-zinc-400'}`}><Package size={24} /></button>
                            <button onClick={obtenerCarrito} className={`${vista === 'carrito' ? 'text-emerald-400' : 'text-zinc-400'}`}><ShoppingCart size={24} /></button>
                            <button onClick={obtenerPerfil} className={`${vista === 'perfil' ? 'text-emerald-400' : 'text-zinc-400'}`}><User size={24} /></button>
                            <button onClick={obtenerAuditoria} className={`${vista === 'auditoria' ? 'text-emerald-400' : 'text-zinc-400'}`}><ClockCounterClockwise size={24} /></button>
                        </div>
                        <div className="w-px h-6 bg-zinc-800 hidden md:block"></div>
                        <button onClick={handleLogout} className="text-zinc-400 hover:text-zinc-100 transition-colors" aria-label="Salir">
                            <SignOut size={24} />
                        </button>
                    </div>
                </div>
            </header>

            {/* Main Content */}
            <main className="max-w-7xl mx-auto px-6 py-12 md:py-20">
                
                {/* View: Catálogo */}
                {vista === 'catalogo' && (
                    <div className="space-y-12">
                        <header className="max-w-2xl">
                            <h2 className="text-4xl md:text-5xl font-medium tracking-tight mb-4">Catálogo</h2>
                            <p className="text-zinc-400 text-lg leading-relaxed">Selección de vinos y destilados de nuestra bodega principal. Añade los productos que desees a tu orden.</p>
                        </header>

                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                            {productos && productos.length > 0 ? productos.map((prod, index) => {
                                const id = prod.identificador || prod.id;
                                const stock = prod.inventario !== undefined && prod.inventario !== null ? prod.inventario : 0;
                                const agotado = stock <= 0;
                                const enCarrito = carrito?.articulos
                                    ?.filter(item => item.idProducto === id)
                                    ?.reduce((sum, item) => sum + item.cantidad, 0) || 0;
                                const sinStockAdicional = enCarrito >= stock;

                                return (
                                    <article key={index} className="group relative flex flex-col bg-zinc-900/50 border border-zinc-800/50 rounded-2xl overflow-hidden transition-all hover:bg-zinc-900">
                                        <div className="aspect-square bg-zinc-800 overflow-hidden relative">
                                            <img 
                                                src={prod.imagen || 'https://images.unsplash.com/photo-1569528020524-7489cb0f56a5?auto=format&fit=crop&q=80&w=800'} 
                                                alt={prod.nombre} 
                                                className={`w-full h-full object-cover transition-transform duration-500 group-hover:scale-105 ${agotado ? 'opacity-40 grayscale' : 'opacity-80 group-hover:opacity-100'}`} 
                                            />
                                            <div className="absolute top-4 right-4 bg-zinc-950/80 backdrop-blur px-3 py-1 rounded-full border border-zinc-700/50">
                                                <span className="text-sm font-mono font-medium">${prod.precio || 0}</span>
                                            </div>

                                            {/* Badge de Stock */}
                                            <div className="absolute top-4 left-4">
                                                {agotado ? (
                                                    <span className="bg-rose-500 text-white text-[11px] font-semibold tracking-wider px-2.5 py-1 rounded-full shadow-lg flex items-center gap-1 uppercase">
                                                        Agotado
                                                    </span>
                                                ) : (
                                                    <span className={`text-[11px] font-semibold px-2.5 py-1 rounded-full shadow-lg backdrop-blur flex items-center gap-1 ${
                                                        stock <= 5 
                                                            ? 'bg-amber-500 text-zinc-950 font-bold' 
                                                            : 'bg-zinc-950/80 text-emerald-400 border border-emerald-500/30'
                                                    }`}>
                                                        Stock: {stock}
                                                    </span>
                                                )}
                                            </div>
                                        </div>
                                        <div className="p-6 flex flex-col flex-grow">
                                            <h3 className="text-lg font-medium mb-1 truncate" title={prod.nombre || `Producto ${id}`}>
                                                {prod.nombre || `Producto ${id}`}
                                            </h3>
                                            <div className="flex items-center justify-between text-zinc-500 text-sm mb-6 font-mono">
                                                <span className="uppercase tracking-widest text-xs">Id: {id}</span>
                                                {enCarrito > 0 && (
                                                    <span className="text-xs text-emerald-400 font-sans font-medium">
                                                        En orden: {enCarrito}
                                                    </span>
                                                )}
                                            </div>
                                            <div className="mt-auto">
                                                <button 
                                                    disabled={agotado || sinStockAdicional}
                                                    onClick={() => agregarAlCarrito(prod)}
                                                    className={`w-full flex items-center justify-center gap-2 font-medium px-4 py-3 rounded-xl transition-all ${
                                                        agotado
                                                            ? 'bg-zinc-800/80 text-zinc-500 cursor-not-allowed border border-zinc-800'
                                                            : sinStockAdicional
                                                                ? 'bg-amber-500/10 text-amber-400 border border-amber-500/20 cursor-not-allowed'
                                                                : 'bg-zinc-100 hover:bg-white text-zinc-950 active:scale-[0.98] cursor-pointer'
                                                    }`}
                                                >
                                                    {agotado 
                                                        ? 'Sin Stock' 
                                                        : sinStockAdicional 
                                                            ? 'Tope alcanzado' 
                                                            : <>Agregar <ShoppingCart size={18} /></>}
                                                </button>
                                            </div>
                                        </div>
                                    </article>
                                );
                            }) : (
                                <div className="col-span-full py-20 text-center border border-dashed border-zinc-800 rounded-2xl">
                                    <p className="text-zinc-500">No hay productos disponibles en este momento.</p>
                                </div>
                            )}
                        </div>
                    </div>
                )}

                {/* View: Carrito */}
                {vista === 'carrito' && (
                    <div className="space-y-12 max-w-4xl mx-auto">
                        <header className="flex flex-col md:flex-row md:items-end justify-between gap-6 border-b border-zinc-800/50 pb-8">
                            <div>
                                <h2 className="text-4xl md:text-5xl font-medium tracking-tight mb-4">Tu Orden</h2>
                                <p className="text-zinc-400 text-lg">Revisa los productos antes de finalizar la compra.</p>
                            </div>
                            {carrito?.articulos?.length > 0 && (
                                <button 
                                    onClick={limpiarCarrito}
                                    className="flex items-center gap-2 text-zinc-400 hover:text-red-400 transition-colors text-sm font-medium"
                                >
                                    <Trash size={18} /> Vaciar orden
                                </button>
                            )}
                        </header>

                        {carrito && carrito.articulos && carrito.articulos.length > 0 ? (
                            <div className="space-y-8">
                                <div className="divide-y divide-zinc-800/50">
                                    {carrito.articulos.reduce((acumulador, item) => {
                                        const existente = acumulador.find(i => i.idProducto === item.idProducto);
                                        if (existente) {
                                            existente.cantidad += item.cantidad;
                                        } else {
                                            acumulador.push({ ...item });
                                        }
                                        return acumulador;
                                    }, []).map((item, index) => {
                                        const productoInfo = productos.find(p => (p.identificador || p.id) === item.idProducto);
                                        const nombre = productoInfo ? productoInfo.nombre : `Producto ID: ${item.idProducto}`;
                                        const precio = productoInfo ? productoInfo.precio : 0;
                                        const subtotal = precio * item.cantidad;
                                        const stockBodega = productoInfo ? (productoInfo.inventario ?? 0) : 0;
                                        const superaStock = item.cantidad > stockBodega;

                                        return (
                                            <div key={index} className="flex flex-col md:flex-row md:items-center justify-between py-6 gap-4">
                                                <div className="flex items-center gap-4">
                                                    <div className="w-16 h-16 bg-zinc-900 rounded-lg overflow-hidden flex-shrink-0 relative">
                                                        <img 
                                                            src={productoInfo?.imagen || 'https://images.unsplash.com/photo-1569528020524-7489cb0f56a5?auto=format&fit=crop&q=80&w=200'} 
                                                            alt={nombre} 
                                                            className={`w-full h-full object-cover ${superaStock ? 'opacity-40 grayscale' : 'opacity-80'}`}
                                                        />
                                                    </div>
                                                    <div>
                                                        <h4 className="text-lg font-medium">{nombre}</h4>
                                                        <div className="flex items-center gap-3 mt-1">
                                                            <p className="text-zinc-500 text-sm font-mono">${precio} x {item.cantidad}</p>
                                                            <span className="text-xs text-zinc-500">|</span>
                                                            <span className={`text-xs ${superaStock ? 'text-rose-400 font-semibold' : 'text-zinc-400'}`}>
                                                                Stock disponible: {stockBodega}
                                                            </span>
                                                        </div>
                                                        {superaStock && (
                                                            <p className="text-xs text-rose-400 font-medium mt-1">
                                                                ⚠️ Supera las {stockBodega} unidades disponibles en bodega
                                                            </p>
                                                        )}
                                                    </div>
                                                </div>
                                                <div className="text-right">
                                                    <span className="text-xl font-medium text-emerald-400">${subtotal}</span>
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                                
                                <div className="bg-zinc-900/50 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row md:items-center justify-between gap-6 border border-zinc-800/50">
                                    <div>
                                        <p className="text-zinc-400 text-sm mb-1">Total a pagar</p>
                                        <p className="text-3xl font-medium text-emerald-400">
                                            ${carrito.articulos.reduce((total, item) => {
                                                const prod = productos.find(p => (p.identificador || p.id) === item.idProducto);
                                                return total + (prod ? prod.precio * item.cantidad : 0);
                                            }, 0)}
                                        </p>
                                    </div>
                                    <button 
                                        disabled={carrito.articulos.some(item => {
                                            const p = productos.find(x => (x.identificador || x.id) === item.idProducto);
                                            return item.cantidad > (p ? (p.inventario ?? 0) : 0);
                                        })}
                                        onClick={procederAlPago} 
                                        className={`flex items-center justify-center gap-2 font-medium px-8 py-4 rounded-xl transition-all w-full md:w-auto ${
                                            carrito.articulos.some(item => {
                                                const p = productos.find(x => (x.identificador || x.id) === item.idProducto);
                                                return item.cantidad > (p ? (p.inventario ?? 0) : 0);
                                            })
                                                ? 'bg-zinc-800 text-zinc-500 cursor-not-allowed border border-zinc-700/50'
                                                : 'bg-emerald-500 hover:bg-emerald-600 text-zinc-950 active:scale-[0.98] cursor-pointer'
                                        }`}
                                    >
                                        {carrito.articulos.some(item => {
                                            const p = productos.find(x => (x.identificador || x.id) === item.idProducto);
                                            return item.cantidad > (p ? (p.inventario ?? 0) : 0);
                                        }) ? 'Stock insuficiente' : 'Proceder al pago'} <ArrowRight size={20} />
                                    </button>
                                </div>
                            </div>
                        ) : (
                            <div className="py-20 text-center border border-dashed border-zinc-800 rounded-2xl">
                                <ShoppingCart size={48} className="mx-auto text-zinc-600 mb-4" />
                                <p className="text-zinc-400 text-lg">Tu orden está vacía.</p>
                                <button 
                                    onClick={obtenerProductos}
                                    className="mt-6 text-emerald-400 hover:text-emerald-300 font-medium"
                                >
                                    Volver al catálogo &rarr;
                                </button>
                            </div>
                        )}
                    </div>
                )}

                {/* View: Perfil */}
                {vista === 'perfil' && (
                    <div className="max-w-2xl mx-auto space-y-12">
                        <header>
                            <h2 className="text-4xl md:text-5xl font-medium tracking-tight mb-4">Perfil</h2>
                            <p className="text-zinc-400 text-lg">Información de tu cuenta.</p>
                        </header>

                        {perfil ? (
                            <div className="bg-zinc-900/50 border border-zinc-800/50 rounded-2xl p-8 space-y-6">
                                <div className="flex items-center gap-6 pb-6 border-b border-zinc-800/50">
                                    <div className="w-20 h-20 bg-zinc-800 rounded-full flex items-center justify-center text-zinc-500">
                                        <User size={32} />
                                    </div>
                                    <div>
                                        <h3 className="text-2xl font-medium">{perfil.nombreUsuario}</h3>
                                        <p className="text-zinc-500 font-mono text-sm mt-1">ID: {perfil.identificador}</p>
                                    </div>
                                </div>
                                <div className="space-y-4">
                                    <div>
                                        <p className="text-sm text-zinc-500 mb-1">Usuario de Azure AD</p>
                                        <p className="font-medium">{perfil.nombreUsuario}</p>
                                    </div>
                                    <div>
                                        <p className="text-sm text-zinc-500 mb-1">Estado de la cuenta</p>
                                        <div className="flex items-center gap-2">
                                            <div className="w-2 h-2 rounded-full bg-emerald-500"></div>
                                            <span className="font-medium text-emerald-400">Activa</span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        ) : (
                            <div className="py-20 text-center border border-dashed border-zinc-800 rounded-2xl">
                                <p className="text-zinc-500">No se pudo cargar la información del perfil para {accounts[0]?.username}.</p>
                            </div>
                        )}
                    </div>
                )}

                {/* View: Auditoría y Trazabilidad */}
                {vista === 'auditoria' && (() => {
                    const logsFiltrados = auditoriaLogs.filter(item => 
                        filtroAuditoria === 'TODOS' || (item.origen && item.origen.toUpperCase() === filtroAuditoria)
                    );
                    const totalPaginas = Math.ceil(logsFiltrados.length / logsPorPagina) || 1;
                    const paginaSegura = Math.max(1, Math.min(paginaAuditoria, totalPaginas));
                    const inicio = (paginaSegura - 1) * logsPorPagina;
                    const logsPaginados = logsFiltrados.slice(inicio, inicio + logsPorPagina);

                    return (
                        <div className="space-y-8">
                            <header className="flex flex-col md:flex-row md:items-end justify-between gap-4">
                                <div>
                                    <div className="flex items-center gap-2 text-emerald-400 text-sm font-medium mb-2">
                                        <ShieldCheck size={20} weight="fill" />
                                        <span>TRAZABILIDAD Y EVENTOS ASÍNCRONOS</span>
                                    </div>
                                    <h2 className="text-4xl md:text-5xl font-medium tracking-tight">Auditoría del Sistema</h2>
                                    <p className="text-zinc-400 text-lg mt-2 max-w-2xl">
                                        Trazabilidad en tiempo real de acciones (agregar/eliminar) y flujo de los microservicios vinculados con RabbitMQ y Supabase.
                                    </p>
                                </div>
                                <div className="flex items-center gap-3 relative">
                                    <button 
                                        onClick={obtenerAuditoria}
                                        className="flex items-center gap-2 px-4 py-2.5 rounded-xl bg-zinc-900 hover:bg-zinc-800 border border-zinc-800 text-zinc-200 text-sm font-medium transition-all active:scale-95 cursor-pointer"
                                    >
                                        <ArrowClockwise size={18} />
                                        Actualizar
                                    </button>

                                    {/* Botón de Limpiar Auditorías con opciones */}
                                    <div className="relative">
                                        <button 
                                            onClick={() => setMenuLimpiarAbierto(prev => !prev)}
                                            className="flex items-center gap-2 px-4 py-2.5 rounded-xl bg-rose-500/10 hover:bg-rose-500/20 border border-rose-500/30 text-rose-300 text-sm font-medium transition-all active:scale-95 cursor-pointer"
                                        >
                                            <Broom size={18} />
                                            Limpiar logs
                                        </button>

                                        {menuLimpiarAbierto && (
                                            <div className="absolute right-0 mt-2 w-64 bg-zinc-900 border border-zinc-700/80 rounded-2xl p-2 shadow-2xl z-50 backdrop-blur-xl">
                                                <div className="px-3 py-2 text-[11px] font-semibold text-zinc-400 uppercase tracking-wider border-b border-zinc-800">
                                                    Opciones de limpieza
                                                </div>
                                                <button
                                                    onClick={() => limpiarAuditoria('10')}
                                                    className="w-full text-left px-3 py-2.5 text-xs text-zinc-200 hover:bg-zinc-800/80 rounded-xl transition-colors flex items-center justify-between"
                                                >
                                                    <span>Limpiar últimos 10 logs</span>
                                                    <span className="text-[10px] bg-zinc-800 px-1.5 py-0.5 rounded text-zinc-400 font-mono">10</span>
                                                </button>
                                                <button
                                                    onClick={() => limpiarAuditoria('pagina')}
                                                    className="w-full text-left px-3 py-2.5 text-xs text-zinc-200 hover:bg-zinc-800/80 rounded-xl transition-colors flex items-center justify-between"
                                                >
                                                    <span>Limpiar esta página ({logsPaginados.length})</span>
                                                    <span className="text-[10px] bg-zinc-800 px-1.5 py-0.5 rounded text-zinc-400 font-mono">{logsPaginados.length}</span>
                                                </button>
                                                <div className="h-px bg-zinc-800 my-1"></div>
                                                <button
                                                    onClick={() => limpiarAuditoria('todos')}
                                                    className="w-full text-left px-3 py-2.5 text-xs text-rose-400 hover:bg-rose-500/10 rounded-xl transition-colors font-medium flex items-center justify-between"
                                                >
                                                    <span>Limpiar todos los logs</span>
                                                    <span className="text-[10px] bg-rose-500/20 px-1.5 py-0.5 rounded text-rose-300 font-mono">{auditoriaLogs.length}</span>
                                                </button>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </header>

                            {/* Pipeline visual del Flujo de Microservicios */}
                            <div className="bg-zinc-900/60 border border-zinc-800/70 rounded-2xl p-6 md:p-8 backdrop-blur-md">
                                <h3 className="text-sm font-semibold uppercase tracking-wider text-zinc-400 mb-6 flex items-center gap-2">
                                    <Lightning size={18} className="text-amber-400" weight="fill" />
                                    Flujo de Trazabilidad en RabbitMQ (botilleria.exchange)
                                </h3>
                                <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
                                    <div className="p-4 rounded-xl bg-zinc-950/80 border border-emerald-500/30">
                                        <span className="text-[11px] font-mono text-emerald-400 font-semibold uppercase">Paso 1: Emisión</span>
                                        <h4 className="font-medium text-emerald-400 mt-1 flex items-center gap-1.5"><Package size={16} /> ms-ordenes</h4>
                                        <p className="text-xs text-zinc-400 mt-1.5">Guarda la compra y emite el evento <code className="text-emerald-300">orden.creada</code></p>
                                    </div>
                                    <div className="p-4 rounded-xl bg-amber-500/10 border border-amber-500/30">
                                        <span className="text-[11px] font-mono text-amber-400 font-semibold uppercase">Broker Asíncrono</span>
                                        <h4 className="font-medium text-amber-400 mt-1 flex items-center gap-1.5"><Lightning size={16} weight="fill" /> RabbitMQ</h4>
                                        <p className="text-xs text-zinc-400 mt-1.5">Distribuye el mensaje a 3 colas en <code className="text-amber-300">34.236.39.87</code></p>
                                    </div>
                                    <div className="p-4 rounded-xl bg-zinc-950/80 border border-sky-500/30">
                                        <span className="text-[11px] font-mono text-sky-400 font-semibold uppercase">Paso 2: Correo</span>
                                        <h4 className="font-medium text-sky-400 mt-1 flex items-center gap-1.5"><EnvelopeSimple size={16} /> ms-notificaciones</h4>
                                        <p className="text-xs text-zinc-400 mt-1.5">Consume <code className="text-sky-300">notificaciones.queue</code> y simula email</p>
                                    </div>
                                    <div className="p-4 rounded-xl bg-zinc-950/80 border border-purple-500/30">
                                        <span className="text-[11px] font-mono text-purple-400 font-semibold uppercase">Paso 3: Logística</span>
                                        <h4 className="font-medium text-purple-400 mt-1 flex items-center gap-1.5"><Truck size={16} /> ms-despachos</h4>
                                        <p className="text-xs text-zinc-400 mt-1.5">Consume <code className="text-purple-300">despachos.queue</code> y genera ticket BTL</p>
                                    </div>
                                    <div className="p-4 rounded-xl bg-zinc-950/80 border border-indigo-500/30">
                                        <span className="text-[11px] font-mono text-indigo-400 font-semibold uppercase">Paso 4: Stock</span>
                                        <h4 className="font-medium text-indigo-400 mt-1 flex items-center gap-1.5"><Database size={16} /> ms-productos</h4>
                                        <p className="text-xs text-zinc-400 mt-1.5">Consume <code className="text-indigo-300">stock.queue</code> y descuenta bodega</p>
                                    </div>
                                </div>
                            </div>

                            {/* Filtros de Logs */}
                            <div className="flex items-center gap-2 overflow-x-auto pb-1">
                                {['TODOS', 'MS-ORDENES', 'MS-NOTIFICACIONES', 'MS-DESPACHOS', 'MS-PRODUCTOS', 'FRONTEND'].map((tag) => {
                                    const totalTag = tag === 'TODOS' 
                                        ? auditoriaLogs.length 
                                        : auditoriaLogs.filter(l => l.origen && l.origen.toUpperCase() === tag).length;

                                    return (
                                        <button
                                            key={tag}
                                            onClick={() => cambiarFiltroAuditoria(tag)}
                                            className={`px-4 py-2 rounded-xl text-xs font-medium transition-all cursor-pointer flex items-center gap-1.5 ${
                                                filtroAuditoria === tag 
                                                    ? 'bg-emerald-500 text-zinc-950 font-semibold shadow-lg shadow-emerald-500/20' 
                                                    : 'bg-zinc-900/80 hover:bg-zinc-800 text-zinc-400 border border-zinc-800/80'
                                            }`}
                                        >
                                            <span>{tag}</span>
                                            <span className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                                                filtroAuditoria === tag ? 'bg-zinc-950/20 text-zinc-950' : 'bg-zinc-800 text-zinc-400'
                                            }`}>
                                                {totalTag}
                                            </span>
                                        </button>
                                    );
                                })}
                            </div>

                            {/* Barra de Paginación con mini pestañas 1, 2, 3... */}
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 py-3 px-4 bg-zinc-900/40 border border-zinc-800/60 rounded-2xl">
                                <div className="text-xs text-zinc-400 font-mono">
                                    Mostrando <span className="text-zinc-200 font-semibold">{logsFiltrados.length === 0 ? 0 : inicio + 1} - {Math.min(inicio + logsPorPagina, logsFiltrados.length)}</span> de <span className="text-emerald-400 font-semibold">{logsFiltrados.length}</span> registros
                                </div>

                                <div className="flex items-center gap-1.5 flex-wrap">
                                    <button
                                        disabled={paginaSegura <= 1}
                                        onClick={() => setPaginaAuditoria(prev => Math.max(1, prev - 1))}
                                        className="p-1.5 rounded-lg bg-zinc-900 border border-zinc-800 text-zinc-400 hover:text-zinc-100 disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer"
                                        title="Página anterior"
                                    >
                                        <CaretLeft size={16} />
                                    </button>

                                    {/* Mini pestañas numeradas */}
                                    {Array.from({ length: totalPaginas }, (_, i) => i + 1).map((numPagina) => (
                                        <button
                                            key={numPagina}
                                            onClick={() => setPaginaAuditoria(numPagina)}
                                            className={`min-w-[30px] h-7 px-2 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
                                                paginaSegura === numPagina
                                                    ? 'bg-emerald-500 text-zinc-950 shadow-md shadow-emerald-500/20'
                                                    : 'bg-zinc-900 hover:bg-zinc-800 text-zinc-400 border border-zinc-800/80 hover:text-zinc-200'
                                            }`}
                                        >
                                            {numPagina}
                                        </button>
                                    ))}

                                    <button
                                        disabled={paginaSegura >= totalPaginas}
                                        onClick={() => setPaginaAuditoria(prev => Math.min(totalPaginas, prev + 1))}
                                        className="p-1.5 rounded-lg bg-zinc-900 border border-zinc-800 text-zinc-400 hover:text-zinc-100 disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer"
                                        title="Página siguiente"
                                    >
                                        <CaretRight size={16} />
                                    </button>

                                    {/* Selector de cantidad por página */}
                                    <div className="ml-3 flex items-center gap-1.5 text-xs text-zinc-400 border-l border-zinc-800 pl-3">
                                        <span>Por página:</span>
                                        <select
                                            value={logsPorPagina}
                                            onChange={(e) => {
                                                setLogsPorPagina(Number(e.target.value));
                                                setPaginaAuditoria(1);
                                            }}
                                            className="bg-zinc-900 border border-zinc-800 text-zinc-200 rounded-lg px-2 py-0.5 text-xs outline-none focus:border-emerald-500 cursor-pointer"
                                        >
                                            <option value={5}>5</option>
                                            <option value={10}>10</option>
                                            <option value={20}>20</option>
                                            <option value={50}>50</option>
                                        </select>
                                    </div>
                                </div>
                            </div>

                            {/* Lista cronológica de logs paginados */}
                            <div className="space-y-3">
                                {logsPaginados.length > 0 ? (
                                    logsPaginados.map((log, index) => (
                                        <div 
                                            key={log.id || index}
                                            className="bg-zinc-900/50 hover:bg-zinc-900/80 border border-zinc-800/60 rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 transition-all"
                                        >
                                            <div className="flex items-start gap-4">
                                                <span className="text-xs font-mono text-zinc-400 py-1 px-2.5 bg-zinc-950 rounded-lg border border-zinc-800 flex-shrink-0">
                                                    {formatearTimestamp(log.timestamp)}
                                                </span>
                                                <div>
                                                    <div className="flex items-center gap-2 flex-wrap mb-1.5">
                                                        <span className={`text-[11px] font-semibold tracking-wider px-2.5 py-0.5 rounded-full border ${
                                                            log.origen === 'MS-ORDENES' ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' :
                                                            log.origen === 'MS-NOTIFICACIONES' ? 'bg-sky-500/10 text-sky-400 border-sky-500/20' :
                                                            log.origen === 'MS-DESPACHOS' ? 'bg-purple-500/10 text-purple-400 border-purple-500/20' :
                                                            log.origen === 'MS-PRODUCTOS' ? 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20' :
                                                            log.origen === 'FRONTEND' ? 'bg-amber-500/10 text-amber-400 border-amber-500/20' :
                                                            'bg-zinc-800 text-zinc-300 border-zinc-700'
                                                        }`}>
                                                            {log.origen || 'SISTEMA'}
                                                        </span>
                                                        <span className="text-xs font-mono text-zinc-400 font-semibold">
                                                            {log.accion}
                                                        </span>
                                                    </div>
                                                    <p className="text-sm text-zinc-200">
                                                        {log.descripcion}
                                                    </p>
                                                </div>
                                            </div>

                                            {log.detalles && Object.keys(log.detalles).length > 0 && (
                                                <div className="text-right flex-shrink-0">
                                                    <span className="text-[11px] font-mono text-zinc-400 bg-zinc-950/80 px-2.5 py-1.5 rounded-lg border border-zinc-800/60 inline-block max-w-[280px] truncate" title={JSON.stringify(log.detalles)}>
                                                        {JSON.stringify(log.detalles)}
                                                    </span>
                                                </div>
                                            )}
                                        </div>
                                    ))
                                ) : (
                                    <div className="py-16 text-center border border-dashed border-zinc-800 rounded-2xl">
                                        <p className="text-zinc-500">No hay registros de auditoría para el filtro seleccionado.</p>
                                        <button 
                                            onClick={() => cambiarFiltroAuditoria('TODOS')}
                                            className="mt-4 text-xs font-medium text-emerald-400 hover:text-emerald-300 cursor-pointer"
                                        >
                                            Ver todos los registros
                                        </button>
                                    </div>
                                )}
                            </div>
                        </div>
                    );
                })()}
            </main>

            {/* Contenedor de Notificaciones Toast apiladas */}
            <div className="fixed bottom-6 right-6 z-50 flex flex-col gap-3 max-w-sm sm:max-w-md w-full pointer-events-none">
                {notificaciones.map((notif) => (
                    <div 
                        key={notif.id}
                        className="pointer-events-auto flex items-center justify-between gap-3 px-5 py-4 rounded-2xl bg-zinc-900/95 border border-zinc-700/60 shadow-2xl backdrop-blur-xl transition-all duration-300"
                    >
                        <div className="flex items-center gap-3 min-w-0">
                            {notif.tipo === 'exito' && (
                                <div className="p-1.5 bg-emerald-500/10 rounded-xl text-emerald-400 flex-shrink-0">
                                    <CheckCircle size={22} weight="fill" />
                                </div>
                            )}
                            {notif.tipo === 'error' && (
                                <div className="p-1.5 bg-rose-500/10 rounded-xl text-rose-400 flex-shrink-0">
                                    <WarningCircle size={22} weight="fill" />
                                </div>
                            )}
                            {notif.tipo === 'info' && (
                                <div className="p-1.5 bg-sky-500/10 rounded-xl text-sky-400 flex-shrink-0">
                                    <Info size={22} weight="fill" />
                                </div>
                            )}
                            <span className="text-sm font-medium text-zinc-100 break-words">{notif.mensaje}</span>
                        </div>
                        <button 
                            onClick={() => cerrarNotificacion(notif.id)}
                            className="p-1.5 text-zinc-500 hover:text-zinc-300 transition-colors rounded-lg hover:bg-zinc-800 flex-shrink-0 ml-2"
                            aria-label="Cerrar notificación"
                        >
                            <X size={16} />
                        </button>
                    </div>
                ))}
            </div>
        </div>
    );
}

export default App;