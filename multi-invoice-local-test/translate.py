# from deep_translator import GoogleTranslator
 
# # Automatically detects source language and translates to Spanish

# translated = GoogleTranslator(source='auto', target='en').translate("El Corte Inglés S.A.") 

# print(translated) # Output: Hola mundo

from vietnam_renamer import (
 QwenTranslator
)
translator = QwenTranslator()
vietnam_translator = QwenTranslator()

result = translator.translate("Công Ty TNHH Intel Products Việt Nam")
print(result)

print('Hello')