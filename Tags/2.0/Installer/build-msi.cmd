SET Wix_Home=".\wix-3.0.4805.0"
SET Path=%Path%;%Wix_Home%

REM Build
candle.exe iResearch.wxs > 1_candle_output.txt
light.exe iResearch.wixobj > 2_light_output.txt


PAUSE