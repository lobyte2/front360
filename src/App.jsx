import React, { useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';

function App() {
    const { instance, accounts } = useMsal();
    const isAuthenticated = useIsAuthenticated();
    const [respuestaApi, setRespuestaApi] = useState(null);

    // Flujo de Autenticación
    const handleLogin = () => {
        instance.loginRedirect(loginRequest);
    };

    const handleLogout = () => {
        instance.logoutRedirect();
    };

    // Llamada protegida al AWS API Gateway usando acquireTokenSilent para ver los logs en la consola principal
    const consultarCarrito = async () => {
        try {
            const tokenResponse = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });

            const token = tokenResponse.accessToken;
            console.log("Token JWT obtenido:", token);

            // Invocamos la URL pública de tu API Gateway con el ID 1
            const response = await fetch("https://ap96rduot1.execute-api.us-east-1.amazonaws.com/api/carrito/1", {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${token}`,
                    "Content-Type": "application/json"
                }
            });

            if (response.ok) {
                const data = await response.json();
                setRespuestaApi(data);
                console.log("¡Conexión Exitosa 200 OK!", data);
            } else {
                console.error("Error devuelto por el Gateway:", response.status);
                const errorText = await response.text();
                console.error("Detalle del error:", errorText);
            }
        } catch (error) {
            console.error("Error de autenticación o de red:", error);
        }
    };

    return (
        <div style={{ padding: "2rem", fontFamily: "sans-serif" }}>
            <h1>Tienda Pedidos360</h1>
            
            {isAuthenticated ? (
                <div>
                    <p>Bienvenido, <strong>{accounts[0]?.name}</strong></p>
                    <button onClick={handleLogout}>Cerrar Sesión</button>
                    
                    <hr style={{ margin: "2rem 0" }} />
                    
                    <h3>Prueba de Integración Frontend {"->"} AWS API Gateway {"->"} EC2</h3>
                    <button onClick={consultarCarrito}>Obtener Datos del Carrito</button>
                    
                    {respuestaApi && (
                        <pre style={{ background: "#f4f4f4", padding: "1rem", marginTop: "1rem" }}>
                            {JSON.stringify(respuestaApi, null, 2)}
                        </pre>
                    )}
                </div>
            ) : (
                <div>
                    <p>Por favor, inicia sesión para acceder al catálogo y carrito.</p>
                    <button onClick={handleLogin}>Iniciar Sesión con Azure AD</button>
                </div>
            )}
        </div>
    );
}

export default App;