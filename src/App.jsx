import React, { useState, useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import { ShoppingCart, User, Wine, SignOut, ArrowRight, Trash, Package } from '@phosphor-icons/react';

const API_GATEWAY_URL = "https://ap96rduot1.execute-api.us-east-1.amazonaws.com/api";
const USUARIO_ID = 1; 

function App() {
    const { instance, accounts } = useMsal();
    const isAuthenticated = useIsAuthenticated();
    
    const [vista, setVista] = useState('catalogo');
    const [productos, setProductos] = useState([]);
    const [carrito, setCarrito] = useState(null);
    const [perfil, setPerfil] = useState(null);

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
        try {
            await peticionAutenticada(`/carrito/${USUARIO_ID}/agregar`, {
                method: 'POST',
                body: JSON.stringify({
                    idProducto: articulo.identificador || articulo.id, 
                    cantidad: 1 
                })
            });
            alert(`Producto agregado al carrito 🍷`);
        } catch (error) {
            console.error("Error al agregar producto:", error);
        }
    };

    const limpiarCarrito = async () => {
        try {
            await peticionAutenticada(`/carrito/${USUARIO_ID}/limpiar`, { method: 'DELETE' });
            alert("Carrito limpiado correctamente.");
            obtenerCarrito();
        } catch (error) {
            console.error("Error al limpiar carrito:", error);
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
                    </nav>

                    <div className="flex items-center gap-4">
                        <div className="md:hidden flex gap-4">
                            <button onClick={obtenerProductos} className={`${vista === 'catalogo' ? 'text-emerald-400' : 'text-zinc-400'}`}><Package size={24} /></button>
                            <button onClick={obtenerCarrito} className={`${vista === 'carrito' ? 'text-emerald-400' : 'text-zinc-400'}`}><ShoppingCart size={24} /></button>
                            <button onClick={obtenerPerfil} className={`${vista === 'perfil' ? 'text-emerald-400' : 'text-zinc-400'}`}><User size={24} /></button>
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
                            {productos && productos.length > 0 ? productos.map((prod, index) => (
                                <article key={index} className="group relative flex flex-col bg-zinc-900/50 border border-zinc-800/50 rounded-2xl overflow-hidden transition-all hover:bg-zinc-900">
                                    <div className="aspect-square bg-zinc-800 overflow-hidden relative">
                                        <img 
                                            src={prod.imagen || 'https://images.unsplash.com/photo-1569528020524-7489cb0f56a5?auto=format&fit=crop&q=80&w=800'} 
                                            alt={prod.nombre} 
                                            className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-105 opacity-80 group-hover:opacity-100" 
                                        />
                                        <div className="absolute top-4 right-4 bg-zinc-950/80 backdrop-blur px-3 py-1 rounded-full border border-zinc-700/50">
                                            <span className="text-sm font-mono font-medium">${prod.precio || 0}</span>
                                        </div>
                                    </div>
                                    <div className="p-6 flex flex-col flex-grow">
                                        <h3 className="text-lg font-medium mb-1 truncate" title={prod.nombre || `Producto ${prod.identificador || prod.id}`}>
                                            {prod.nombre || `Producto ${prod.identificador || prod.id}`}
                                        </h3>
                                        <p className="text-zinc-500 text-sm mb-6 font-mono uppercase tracking-widest">Id: {prod.identificador || prod.id}</p>
                                        <div className="mt-auto">
                                            <button 
                                                onClick={() => agregarAlCarrito(prod)}
                                                className="w-full flex items-center justify-center gap-2 bg-zinc-100 hover:bg-white text-zinc-950 font-medium px-4 py-3 rounded-xl transition-all active:scale-[0.98]"
                                            >
                                                Agregar <ShoppingCart size={18} />
                                            </button>
                                        </div>
                                    </div>
                                </article>
                            )) : (
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

                                        return (
                                            <div key={index} className="flex flex-col md:flex-row md:items-center justify-between py-6 gap-4">
                                                <div className="flex items-center gap-4">
                                                    <div className="w-16 h-16 bg-zinc-900 rounded-lg overflow-hidden flex-shrink-0">
                                                        <img 
                                                            src={productoInfo?.imagen || 'https://images.unsplash.com/photo-1569528020524-7489cb0f56a5?auto=format&fit=crop&q=80&w=200'} 
                                                            alt={nombre}
                                                            className="w-full h-full object-cover opacity-80"
                                                        />
                                                    </div>
                                                    <div>
                                                        <h4 className="text-lg font-medium">{nombre}</h4>
                                                        <p className="text-zinc-500 text-sm font-mono">${precio} x {item.cantidad}</p>
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
                                    <button className="flex items-center justify-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-zinc-950 font-medium px-8 py-4 rounded-xl transition-all active:scale-[0.98] w-full md:w-auto">
                                        Proceder al pago <ArrowRight size={20} />
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
            </main>
        </div>
    );
}

export default App;