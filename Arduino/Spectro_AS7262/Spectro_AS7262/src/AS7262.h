/**
   AS7262.h Library to interact with AS7262 sensor.
   Creted by Crhistian Segura
   Change Log:
      Date    |   Version   | Description
   31.01.2019 |    1.0.0    | Creation Lib.
**/
#ifndef AS7262_H
#define AS7262_H
#include "Arduino.h"

enum configParam {
  UART = 1 << 0,
  I2C = 1 << 1
};

class AS7262 {
  public:
    AS7262();
    void init(configParam param = 0);
  private:
    int _address = 0x49;
};

#endif //AS7262_H
