"""Diagnóstico isolado dos dois componentes usados na calibração:
1) detecção da janela ativa (pygetwindow)
2) captura da tecla F8 (keyboard)

Rode: python debug_calibration.py
Siga as instruções impressas no terminal.
"""
import time

import keyboard
import pygetwindow as gw
import pyautogui

print("=== TESTE 1: listar todas as janelas visíveis ===")
for w in gw.getAllWindows():
    if w.title.strip():
        print(f"  - '{w.title}'  pos=({w.left},{w.top})  tamanho=({w.width}x{w.height})  visible={w.visible}")

print("\n=== TESTE 2: janela ativa ===")
print("Clique agora na janela do Citrix (deixe ela em foco) e volte pra cá em 3 segundos...")
time.sleep(3)
active = gw.getActiveWindow()
if active is None:
    print("  FALHOU: gw.getActiveWindow() retornou None.")
else:
    print(f"  OK: janela ativa = '{active.title}' pos=({active.left},{active.top})")

print("\n=== TESTE 3: captura de tecla F8 ===")
print("Pressione F8 em qualquer lugar nos próximos 15 segundos...")
start = time.time()
detected = False
while time.time() - start < 15:
    if keyboard.is_pressed("f8"):
        print(f"  OK: F8 detectado! Posição do mouse no momento: {pyautogui.position()}")
        detected = True
        break
    time.sleep(0.02)
if not detected:
    print("  FALHOU: F8 não foi detectado em 15 segundos.")

print("\n=== TESTE 4: captura de tecla ESC (mesmo mecanismo) ===")
print("Pressione ESC nos próximos 10 segundos...")
start = time.time()
detected = False
while time.time() - start < 10:
    if keyboard.is_pressed("esc"):
        print("  OK: ESC detectado!")
        detected = True
        break
    time.sleep(0.02)
if not detected:
    print("  FALHOU: ESC não foi detectado.")

print("\nFim do diagnóstico.")
