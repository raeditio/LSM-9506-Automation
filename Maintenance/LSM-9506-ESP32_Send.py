#include <WiFi.h>
#include <WiFiClient.h>

const char* ssid = "WescanRichmond";
const char* password = "PYL0Ne1ab";

const char* serverIP = "192.168.1.100"; // Replace with your computer's IP address
const int serverPort = 5555; // Choose a port number

WiFiClient client;

void setup() {
  Serial.begin(115200);
  delay(1000);

  Serial.println("Connecting to WiFi...");
  WiFi.begin(ssid, password);
  
  while (WiFi.status() != WL_CONNECTED) {
    delay(1000);
    Serial.println("Connecting...");
  }
  Serial.println("Connected to WiFi");
}

void loop() {
  if (!client.connected()) {
    Serial.println("Connecting to server...");
    if (client.connect(serverIP, serverPort)) {
      Serial.println("Connected to server");
      client.println("Hello from ESP32"); // Replace with the data you want to send
    } else {
      Serial.println("Connection failed");
    }
  }

  delay(5000); // Adjust the delay as needed
}
