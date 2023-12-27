import socket

HOST = '0.0.0.0'  # Accept connections on any available network interface
PORT = 8888        # Use the same port as defined in your ESP32 code

with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server_socket:
    server_socket.bind((HOST, PORT))
    server_socket.listen()

    print(f"Server listening on port {PORT}")

    conn, addr = server_socket.accept()
    with conn:
        print('Connected by', addr)
        while True:
            data = conn.recv(1024)  # Receive data (adjust buffer size as needed)
            if not data:
                break
            print('Received:', data.decode())
            # Here, you can write the received data to a file, database, or perform any required processing