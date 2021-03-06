#include "M5Atom.h"

#include "Save.h"
#include "WPS.h"
#include "Server.h"
#include "Spreadsheet.h"
#include "Sensor.h"

const int BUTTON_WPS = 39;
const int TEMPERATURE_OFFSET = 4;

struct SaveData saveData;
unsigned long oldTime;
Measurements measurements;
const int hueRed = 98;
float co2LED = 1;

void setup()
{
  M5.begin(true, false, true); // Serial, I2C, LED
  setCpuFrequencyMhz(60);
  pinMode(BUTTON_WPS, INPUT);
  Wire.begin(25, 21);
  delay(50);
  //EEPROM
  M5.dis.drawpix(0, CRGB::Magenta);
  saveSetup();
  dataGet(saveData);
  //Sensor
  M5.dis.drawpix(0, CRGB::Yellow);
  sensorSetup(TEMPERATURE_OFFSET);
  //WPS
  M5.dis.drawpix(0, CRGB::Cyan);
  WPSSetup();
  //Server
  M5.dis.drawpix(0, CRGB::Green);
  serverSetup();
  serverSet(String(saveData.uri));
  //Time
  oldTime =  millis();
}
void loop() {
  unsigned long nowTime = millis();
  //LED
  float brightness = sin(nowTime / 2000.0) * 0.3 + 0.7;
  M5.dis.drawpix(0, CHSV(hueRed * co2LED, 255, 255 * brightness));
  //Measurement and push
  if (abs(nowTime - oldTime) > 600000) {
    oldTime = nowTime;
    M5.dis.drawpix(0, CRGB::Blue);
    sensorGet(measurements);
    String dataString = String(measurements.co2) + " " + String(measurements.temperature) + " " + String(measurements.humidity);
    Serial.println(dataString);
    spreadsheetPut(saveData.uri, dataString);
    co2LED = constrain((measurements.co2 - 400) / 1600.0, 0, 1);
  }
  //Server standby
  if (!digitalRead(BUTTON_WPS)) {
    WPSStart();
  }
  if (serverLoop()) {
    Serial.print("URI Update:");
    M5.dis.drawpix(0, CRGB::White);
    serverGet().toCharArray(saveData.uri, 256);
    Serial.println(saveData.uri);
    dataSave(saveData);
  }
  //M5
  M5.update();
}
