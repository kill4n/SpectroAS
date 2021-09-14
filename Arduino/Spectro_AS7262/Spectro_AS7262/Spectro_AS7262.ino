/**
   Spectro Pothometer with Arduino and AS7262
   Creted by Crhistian Segura
   Change Log:
      Date    |   Version   | Description
   31.01.2019 |    2.0.0    | Creation AS7262.
   11.02.2019 |    2.0.1    | Add Leds ON/OFF RGB
   ********************************************************************
   Default Pinout I2C
   (Sensor -> Arduino uno)
   -----------------------
   SDA -> SDA (A4)
   SCL -> SCL (A5)

 **/

// Incluir libreria de sparkfun AS726X
#include "AS726X.h"
AS726X sensor;//Creates the sensor object

//RGBW Leds Pinouts
#define LEDR 11
#define LEDG 10
#define LEDB 9
//#define LEDW 7

void setup() {
  // Leds direction
  pinMode(LEDR, OUTPUT); pinMode(LEDG, OUTPUT); pinMode(LEDB, OUTPUT);
  Serial.begin(115200);
  sensor.begin(Wire, 0, 3);//Initializes the sensor with non default values
  sensor.setIntegrationTime(10); //10 * 2.8ms = 28ms. 0 to 255 is valid.
  //If you use Mode 2 or 3 (all the colors) then integration time is double. 28*2 = 56ms between readings.
  //sensor.enableIndicator();
}

#define RES 2
enum AS7262Enum {
  MMODE_0 = 0X00, MMODE_1 = 0X01, MMODE_2 = 0X02, MMODE_3 = 0X03, LEDI_ON = 0X10, LEDI_OF = 0X11, LEDI_C1 = 0X12,
  LEDI_C2 = 0X13, LEDI_C3 = 0X14, LEDI_C4 = 0X15, LEDB_ON = 0X20, LEDB_OF = 0X21, LEDB_C1 = 0X22, LEDB_C2 = 0X23,
  LEDB_C3 = 0X24, LEDB_C4 = 0X25, RESET_S = 0X30, GAIN_01 = 0X40, GAIN_02 = 0X41, GAIN_03 = 0X42, GAIN_04 = 0X43,
  TAKE_MS = 0X50, TAKE_PR = 0X60, INTTIME = 0X70, R_ON = 0xC0, R_OF = 0xC1, G_ON = 0xC2, G_OF = 0xC3, B_ON = 0xC4,
  B_OF = 0xC5
};
byte data;
void loop() {
  if (Serial.available() > 0) {
    data=Serial.read();
    switch (data) {
      case MMODE_0: sensor.setMeasurementMode(0); break;
      case MMODE_1: sensor.setMeasurementMode(1); break;
      case MMODE_2: sensor.setMeasurementMode(2); break;
      case MMODE_3: sensor.setMeasurementMode(3); break;
      case LEDI_ON: sensor.enableIndicator(); break;
      case LEDI_OF: sensor.disableIndicator(); break;
      case LEDI_C1: sensor.setIndicatorCurrent(0); break;
      case LEDI_C2: sensor.setIndicatorCurrent(1); break;
      case LEDI_C3: sensor.setIndicatorCurrent(2); break;
      case LEDI_C4: sensor.setIndicatorCurrent(3); break;
      case LEDB_ON: sensor.enableBulb(); break;
      case LEDB_OF: sensor.disableBulb(); break;
      case LEDB_C1: sensor.setBulbCurrent(0); break;
      case LEDB_C2: sensor.setBulbCurrent(1); break;
      case LEDB_C3: sensor.setBulbCurrent(2); break;
      case LEDB_C4: sensor.setBulbCurrent(3); break;
      case RESET_S: sensor.softReset(); break;
      case GAIN_01: sensor.setGain(0); break;
      case GAIN_02: sensor.setGain(1); break;
      case GAIN_03: sensor.setGain(2); break;
      case GAIN_04: sensor.setGain(3); break;
      case TAKE_MS: sensor.takeMeasurements(); break;
      case TAKE_PR:
        sensor.takeMeasurements();
        Serial.print(sensor.getCalibratedViolet(), RES); Serial.print(";");
        Serial.print(sensor.getCalibratedBlue(), RES); Serial.print(";");
        Serial.print(sensor.getCalibratedGreen(), RES); Serial.print(";");
        Serial.print(sensor.getCalibratedYellow(), RES); Serial.print(";");
        Serial.print(sensor.getCalibratedOrange(), RES); Serial.print(";");
        Serial.print(sensor.getCalibratedRed(), RES); Serial.println("");
        break;
      case INTTIME: sensor.setIntegrationTime(Serial.read()); break;
      case R_ON: digitalWrite(LEDR, HIGH); break;
      case R_OF: digitalWrite(LEDR, LOW); break;
      case G_ON: digitalWrite(LEDG, HIGH); break;
      case G_OF: digitalWrite(LEDG, LOW); break;
      case B_ON: digitalWrite(LEDB, HIGH); break;
      case B_OF: digitalWrite(LEDB, LOW); break;
      //case : digitalWrite(LEDW, HIGH); break;
      //case : digitalWrite(LEDW, LOW); break;
      default:
        //Serial.println(Serial.read(), HEX);
        break;
    }
    /*
        //String s = Serial.readStringUntil('/r'); // Until CR (Carriage Return)
      //Serial.println(s);

        switch (s) {
       case "M0": Serial.print("Lei 1:"); Serial.println(s); break;
       case "M1": Serial.print("Lei 2:"); Serial.println(s); break;
       case "M2": Serial.print("Lei 3:"); Serial.println(s); break;
       default: Serial.print("Lei :"); Serial.println(s); break;
       }
         switch (Serial.read())
         {
           case '@':
             //Serial.println(" V | B | G | Y | O | R ");
             sensor.takeMeasurements();
             Serial.print(sensor.getCalibratedViolet(), RES);
             Serial.print(";");
             Serial.print(sensor.getCalibratedBlue(), RES);
             Serial.print(";");
             Serial.print(sensor.getCalibratedGreen(), RES);
             Serial.print(";");
             Serial.print(sensor.getCalibratedYellow(), RES);
             Serial.print(";");
             Serial.print(sensor.getCalibratedOrange(), RES);
             Serial.print(";");
             Serial.print(sensor.getCalibratedRed(), RES);
             Serial.println("");
             break;
           case '$': break;
           default: break;
         }*/
    delay(50);
  }
}
