import React, { useState, useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';

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
            <div style={styles.loginContainer}>
                <h1>🍺 Botillería Pedidos360</h1>
                <p>Inicia sesión para acceder al catálogo.</p>
                <button style={styles.btnPrimary} onClick={handleLogin}>Ingresar con Azure AD</button>
            </div>
        );
    }

    return (
        <div style={styles.appContainer}>
            <nav style={styles.navbar}>
                <h2>🍺 Botillería 360</h2>
                <div style={styles.navLinks}>
                    <button style={styles.btnNav} onClick={obtenerProductos}>Catálogo</button>
                    <button style={styles.btnNav} onClick={obtenerCarrito}>Carrito</button>
                    <button style={styles.btnNav} onClick={obtenerPerfil}>Mi Perfil</button>
                    <button style={styles.btnDanger} onClick={handleLogout}>Salir</button>
                </div>
            </nav>

            <main style={styles.main}>
                {vista === 'catalogo' && (
                    <div style={styles.grid}>
                        {productos && productos.length > 0 ? productos.map((prod, index) => (
                            <div key={index} style={styles.card}>
                                {/* Etiqueta img agregada para mostrar la imagen */}
                                <img 
                                    src={prod.imagen || 'https://via.placeholder.com/200x200?text=Sin+Imagen'} 
                                    alt={prod.nombre} 
                                    style={{ width: '100%', height: '150px', objectFit: 'cover', borderRadius: '5px', marginBottom: '10px' }} 
                                />
                                <h4 style={{ margin: '5px 0' }}>{prod.nombre || `Producto ${prod.identificador || prod.id}`}</h4>
                                <p style={styles.precio}>${prod.precio || 0}</p>
                                <button style={styles.btnPrimary} onClick={() => agregarAlCarrito(prod)}>
                                    Agregar al Carrito
                                </button>
                            </div>
                        )) : <p>Cargando catálogo o sin datos quemados...</p>}
                    </div>
                )}

                {vista === 'carrito' && (
                    <div style={styles.card}>
                        <h3>Mi Carrito</h3>
                        {carrito && carrito.articulos && carrito.articulos.length > 0 ? (
                            <>
                                <ul style={{ listStyle: 'none', padding: 0 }}>
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
                                            <li key={index} style={styles.cartItem}>
                                                <span style={{ textAlign: 'left', width: '40%' }}><strong>{nombre}</strong></span>
                                                <span>${precio} x {item.cantidad}</span>
                                                <strong style={{ color: '#27ae60' }}>${subtotal}</strong>
                                            </li>
                                        );
                                    })}
                                </ul>
                                <h3 style={{ textAlign: 'right', color: '#2c3e50', borderTop: '2px solid #eee', paddingTop: '10px' }}>
                                    Total: ${carrito.articulos.reduce((total, item) => {
                                        const prod = productos.find(p => (p.identificador || p.id) === item.idProducto);
                                        return total + (prod ? prod.precio * item.cantidad : 0);
                                    }, 0)}
                                </h3>
                                <button style={styles.btnDanger} onClick={limpiarCarrito}>Vaciar Carrito</button>
                            </>
                        ) : (
                            <p>Tu carrito está vacío.</p>
                        )}
                    </div>
                )}

                {vista === 'perfil' && (
                    <div style={styles.card}>
                        <h3>Perfil de Usuario</h3>
                        {perfil ? (
                            <div style={{ textAlign: 'left', display: 'inline-block', marginTop: '15px' }}>
                                <p><strong>ID de Sistema:</strong> {perfil.identificador}</p>
                                <p><strong>Usuario (Azure):</strong> {perfil.nombreUsuario}</p>
                                <p><strong>Contraseña:</strong> *********</p>
                            </div>
                        ) : (
                            <p>No se encontraron datos quemados para el usuario {accounts[0]?.username} en el backend.</p>
                        )}
                    </div>
                )}
            </main>
        </div>
    );
}

const styles = {
    loginContainer: { textAlign: 'center', marginTop: '100px', fontFamily: 'sans-serif' },
    appContainer: { fontFamily: 'sans-serif', backgroundColor: '#f4f4f4', minHeight: '100vh' },
    navbar: { display: 'flex', justifyContent: 'space-between', padding: '1rem 2rem', backgroundColor: '#2c3e50', color: 'white', alignItems: 'center' },
    navLinks: { display: 'flex', gap: '10px' },
    main: { padding: '2rem', maxWidth: '900px', margin: '0 auto' },
    grid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '20px' },
    card: { backgroundColor: 'white', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 5px rgba(0,0,0,0.1)', textAlign: 'center' },
    precio: { fontSize: '1.2rem', fontWeight: 'bold', color: '#27ae60' },
    cartItem: { display: 'flex', justifyContent: 'space-between', padding: '10px 0', borderBottom: '1px solid #ddd' },
    btnPrimary: { backgroundColor: '#3498db', color: 'white', border: 'none', padding: '10px', borderRadius: '5px', cursor: 'pointer', width: '100%', marginTop: '10px' },
    btnDanger: { backgroundColor: '#e74c3c', color: 'white', border: 'none', padding: '10px', borderRadius: '5px', cursor: 'pointer', marginTop: '10px' },
    btnNav: { backgroundColor: 'transparent', color: 'white', border: '1px solid white', padding: '8px 12px', borderRadius: '5px', cursor: 'pointer' },
};

export default App;