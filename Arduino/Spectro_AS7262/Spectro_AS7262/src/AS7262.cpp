/**
   AS7262.cpp Library to interact with AS7262 sensor.
   Creted by Crhistian Segura
   Change Log:
      Date    |   Version   | Description
   31.01.2019 |    1.0.0    | Creation Lib.
**/

#include "AS7262.h"

AS7262::AS7262() {

}

void AS7262::init(configParam param) {
  if (param & UART)
  {

  }
  if (param & I2C) {

  }
}
