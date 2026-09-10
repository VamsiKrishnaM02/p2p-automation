"""
Run the file to check you are connected to VLM or not
"""

from vietnam_renamer import (
 QwenTranslator
)
translator = QwenTranslator()
vietnam_translator = QwenTranslator()

result = translator.translate("Công Ty TNHH Intel Products Việt Nam")
print(result)

print('Hello')