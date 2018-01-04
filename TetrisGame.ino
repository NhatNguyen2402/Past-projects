// NeoPixel Ring simple sketch (c) 2013 Shae Erisson
// released under the GPLv3 license to match the rest of the AdaFruit NeoPixel library

#include <Adafruit_NeoPixel.h>
#ifdef __AVR__
  #include <avr/power.h>
#endif

// Which pin on the Arduino is connected to the NeoPixels?
// On a Trinket or Gemma we suggest changing this to 1
#define PIN            6

// How many NeoPixels are attached to the Arduino?
#define NUMPIXELS      200

// When we setup the NeoPixel library, we tell it how many pixels, and which pin to use to send signals.
// Note that for older NeoPixel strips you might need to change the third parameter--see the strandtest
// example for more information on possible values.
Adafruit_NeoPixel pixels = Adafruit_NeoPixel(NUMPIXELS, PIN, NEO_GRB + NEO_KHZ800);

int delayval = 300; // delay for half a second
int reference[200] = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 
                      39, 38, 37, 36, 35, 34, 33, 32, 31, 30, 29, 28, 27, 26, 25, 24, 23, 22, 21, 20,
                      40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59,
                      79, 78, 77, 76, 75, 74, 73, 72, 71, 70, 69, 68, 67, 66, 65, 64, 63, 62, 61, 60,
                      80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99,
                      119, 118, 117, 116, 115, 114, 113, 112, 111, 110, 109, 108, 107, 106, 105, 104, 103, 102, 101, 100,
                      120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139,
                      159, 158, 157, 156, 155, 154, 153, 152, 151, 150, 149, 148, 147, 146, 145, 144, 143, 142, 141, 140,
                      160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179,
                      199, 198, 197, 196, 195, 194, 193, 192, 191, 190, 189, 188, 187, 186, 185, 184, 183, 182, 181, 180};
int light[200] = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};

int tblock[4] = {79, 99, -1, 119};
int iblock[4] = {79, 99, 119, 139};
int oblock[4] = {99, -1, 119, -1};
int sblock[4] = {79, 99, -1, -1};
int zblock[4] = {-1, 99, -1, 119};
int jblock[4] = {79, -1, 99, 119};
int lblock[4] = {79, 99, 119, -1};
int current[4] = {-1, -1, -1, -1};

int shape = 0;
int gameOver = 0;
int height = 0;

void setup() {
  // This is for Trinket 5V 16MHz, you can remove these three lines if you are not using a Trinket
#if defined (__AVR_ATtiny85__)
  if (F_CPU == 16000000) clock_prescale_set(clock_div_1);
#endif
  // End of trinket special code

  pixels.begin(); // This initializes the NeoPixel library.
  pixels.setBrightness(2);
  Serial.begin(9600);
}

void loop() {
  
  // For a set of NeoPixels the first NeoPixel is 0, second is 1, all the way up to the count of pixels minus one.
  /*for(int i = 0; i<20;i++) {
    for(int j=i;j<NUMPIXELS;j+=20){
      // pixels.Color takes RGB values, from 0,0,0 up to 255,255,255
      pixels.setPixelColor(reference[j], pixels.Color(0,150,0)); // Moderately bright green color.
    }
    delay(delayval);
    pixels.show(); // This sends the updated pixel color to the hardware.
  }*/
  Serial.write("reached");
  Serial.println();
  while(gameOver == 0) {
    shape = random(7);
    instantCurrent(shape);
    displayShape(shape);
    //while(canDrop(shape) == 1 && height < 20) {
    while(height < 19) {
      Serial.write("loop reached");
      Serial.println();
      slowdrop(shape);
      Serial.write("slowdrop reached");
      Serial.println();
      delay(delayval);
      height++;
    }
    height = 0;
    /*if(light[19] == 1 || light[39] == 1 || light[59] == 1 || light[79] == 1 || light[99] == 1 || 
      light[119] == 1 || light[139] == 1 || light[159] == 1 || light[179] == 1 || light[199] == 1){
        gameOver = 1;    
    }*/
   
    delay(delayval);
  }
  delay(delayval);
  //if(Serial.available() > 0)
  pixels.show();
}

void slowdrop(int shape) {
  dropShape(shape);
  /*for(int i=0;i<4;i++){
    pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
    light[reference[current[i]]] = 0;
  }*/
  displayShape(shape);  
}

int canDrop(int shape) {
  switch(shape) {
    case 0:
      if(light[current[0]+2] == 0 && light[current[1]+2] == 0 && light[current[3]+2] == 0) {
        return 1; 
      }
      return 0;
      break;
    case 1:
      if(light[current[0]+2] == 0 && light[current[1]+2]==0 && light[current[2]+2]==0 && light[current[3]+2]==0) {
        return 1; 
      }
      return 0;
      break;
    case 2:
      if(light[current[0]+2] == 0 && light[current[2]+2]==0) {
        return 1; 
      }
      return 0;
      break;
    case 3:
      if(light[current[0]+2] == 0 && light[current[1]+2]==0 && light[current[3]+2]==0) {
        return 1; 
      }
      return 0;
      break;
    case 4:
      if(light[current[0]+2] == 0 && light[current[1]+2]==0 && light[current[3]+2]==0) {
        return 1; 
      }
      return 0;
      break;
    case 5:
      if(light[current[0]+2] == 0 && light[current[2]+2]==0 && light[current[3]+2]==0) {
        return 1; 
      }
      return 0;
      break;
    case 6:
      if(light[current[0]+2] == 0 && light[current[1]+2]==0 && light[current[2]+2]==0) {
        return 1;
      }
      return 0;
      break;
  }
}
void instantCurrent(int shape) {
  switch(shape) {
    case 0:
      memcpy(&current, tblock, sizeof(tblock) );
      break;
    case 1:
      memcpy(&current, iblock, sizeof(iblock) );
      break;
    case 2:
      memcpy(&current, oblock, sizeof(oblock) );
      break;
    case 3:
      memcpy(&current, sblock, sizeof(sblock) );
      break;
    case 4:
      memcpy(&current, zblock, sizeof(zblock) );
      break;
    case 5:
      memcpy(&current, jblock, sizeof(jblock) );
      break;
    case 6:
      memcpy(&current, lblock, sizeof(lblock) );
      break;
  }
}
void dropShape(int shape) {
  switch(shape) {
    case 0:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[2] == -2) {
          current[2] = 99;
        }
        //light[current[i]] = 1;
      }
      
      break;
    case 1:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        //light[current[i]] = 1;
      }
      break;
    case 2:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[i] == -2) {
          if(i == 1) {
            current[1] = 99;
          }
          if(i == 3) {
            current[3] = 119;
          }
        }
        //light[current[i]] = 1;
      }
      break;
    case 3:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[i] == -2) {
          if(i == 2) {
            current[2] = 99;
          }
          if(i == 3) {
            current[3] = 119;
          }
        }
        //light[current[i]] = 1;
      }
      
      break;
    case 4:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[i] == -2) {
          if(i == 0) {
            current[0] = 79;
          }
          if(i == 2) {
            current[2] = 99;
          }
        }
        //light[current[i]] = 1;
      }
      
      break;
    case 5:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[1] == -2) {
          current[1] = 79;
        }
        //light[current[i]] = 1;
      }
      break;
    case 6:
      for(int i =0;i<4;i++) {
        if(current[i] != -1) {
          pixels.setPixelColor(reference[current[i]], pixels.Color(0, 0, 0));
          light[reference[current[i]]] = 0;
        }
        current[i]--;
        if(current[3] == -2) {
          current[3] = 119;
        }
        //light[current[i]] = 1;
      }
      break;
  }
  //return rand;
}
void displayShape(int shape) {
  int red = 0;
  int green = 0; 
  int blue = 0;
  switch(shape) {
    case 0:
      red = 231;
      green = 94;
      blue = 255;
      break;
    case 1:
      red = 38;
      green = 201;
      blue = 223;
      break;
    case 2:
      red = 243;
      green = 243;
      blue = 15;
      break;
    case 3:
      red = 99;
      green = 237;
      blue = 29;
      break;
    case 4:
      red = 223;
      green = 38;
      blue = 77;
      break;
    case 5:
      red = 0;
      green = 0;
      blue = 204;
      break;
    case 6:
      red = 253;
      green = 95; 
      blue = 0;
      break;
  }
  for(int x = 0; x < 4; x++){
    if(current[x] != -1) {
      pixels.setPixelColor(reference[current[x]], pixels.Color(red, green, blue));
      light[current[x]] = 1;
    }
  }
  pixels.show();
}


