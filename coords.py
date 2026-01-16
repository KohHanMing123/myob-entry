import pyautogui
import tkinter as tk
from tkinter import font

def create_floating_tracker():
    # Create the main window
    root = tk.Tk()
    root.title("Coords")
    
    # Make it always on top and remove window decorations for minimal look
    root.attributes('-topmost', True)
    root.attributes('-alpha', 0.9)  # Slight transparency
    
    # Set window size and position (top-right corner)
    root.geometry('200x80+1700+10')  # Adjust position as needed
    




































    
    # Style
    root.configure(bg='#1a1a2e')
    
    # Create a custom font
    coord_font = font.Font(family='Consolas', size=16, weight='bold')
    label_font = font.Font(family='Consolas', size=10)
    
    # Labels
    title_label = tk.Label(root, text="Mouse Position", bg='#1a1a2e', fg='#eee', font=label_font)
    title_label.pack(pady=(5, 0))
    
    coord_label = tk.Label(root, text="X: 0000  Y: 0000", bg='#1a1a2e', fg='#00ff88', font=coord_font)
    coord_label.pack(pady=5)
    
    hint_label = tk.Label(root, text="Press Ctrl+C in terminal to stop", bg='#1a1a2e', fg='#666', font=('Consolas', 8))
    hint_label.pack()
    
    def update_position():
        x, y = pyautogui.position()
        coord_label.config(text=f"X: {x:>4}  Y: {y:>4}")
        root.after(50, update_position)  # Update every 50ms
    
    # Start updating
    update_position()
    
    # Run the window
    root.mainloop()

if __name__ == "__main__":
    print("Floating coordinate tracker started. Close the window to stop.")
    create_floating_tracker()