int switchPin = 8;
int ledPin = 2;
bool buttonState = LOW;
void setup()
{
  pinMode(switchPin, INPUT);
  pinMode(ledPin, OUTPUT);
}

void loop()
{
  digitalWrite(ledPin, buttonState);
  if(digitalRead(switchPin) == HIGH)
  {
    buttonState = !buttonState;
  }
  
  
}

/*
int switchPin = 8;
int ledPin = 2;
boolean lastButton = LOW;
boolean currentButton = LOW;
boolean ledOn = false;
int ledLevel = 0;


// Makes signal from button clearer
boolean debounce(boolean last)
{
  boolean current = digitalRead(switchPin);
  if (last != current)
  {
    delay(5);
    current = digitalRead(switchPin);
  }
  return current;
}


void setup(){
  pinMode(switchPin, INPUT);
  pinMode(ledPin, OUTPUT);
}



void loop(){
  currentButton = debounce(lastButton);

  if (lastButton == LOW && currentButton == HIGH)
  {
    ledLevel = ledLevel + 51;
  }
  lastButton = currentButton;

  if (ledLevel > 255) ledLevel = 0;
  analogWrite(ledPin, ledLevel);
}

 basic button shit
int switchPin = 8;
int ledPin = 13;
bool buttonState = LOW;
void setup()
{
  pinMode(switchPin, INPUT);
  pinMode(ledPin, OUTPUT);


}

void loop()
{
  if (digitalRead(switchPin) == HIGH)
  {
    digitalWrite(ledPin, HIGH);
  }
  else
  {
    digitalWrite(ledPin, LOW);
  }
}

 * first program
 * 
 * int ledPin = 13;

void setup() {
  // put your setup code here, to run once:
pinMode(ledPin, OUTPUT);
}

void loop() {
  // put your main code here, to run repeatedly:
digitalWrite(ledPin, HIGH);
delay(5000);
digitalWrite(ledPin, LOW);
delay(5000);
}*/
