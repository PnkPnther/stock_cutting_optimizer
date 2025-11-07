import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

pipes = []
bigFont = ("Arial", 20)


def load_optimize():
    fileName = filedialog.askopenfilename(initialdir="C:", title="Open File", filetypes=(("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")))



    if fileName:
        try:

            data = pd.read_excel(fileName)
            quantity = data["Qty"].tolist()
            cutLength = data["Size"].tolist()

            demands = []


            for i in range(len(quantity)):
                pair = [quantity[i], cutLength[i]]
                demands.append(pair)

            standard_length = int(stockPipeLen.get())

            # Step 1: Expand the list into individual pieces
            pieces = []
            for quantity, length in demands:
                for _ in range(quantity):
                    pieces.append(length)

            # Step 2: Sort pieces descending (greedy best fit)
            pieces.sort(reverse=True)

            # Step 3: Allocate to pipes
            global pipes

            for piece in pieces:
                placed = False
                # Try to fit in an existing pipe
                for pipe in pipes:
                    if sum(pipe) + piece <= standard_length:
                        pipe.append(piece)
                        placed = True
                        break
                # If it doesn't fit in any existing pipe, start a new one
                if not placed:
                    pipes.append([piece])

            drawGraph(pipes,standard_length)




        except Exception as e:  # Catch any errors that occur while loading the file
            messagebox.showerror("Error", f"An error occurred while loading the file")

def drawGraph(final_patterns, stock_length):
    global lastFig
    fig, ax = plt.subplots(figsize=(10, len(final_patterns)))
    lastFig = fig

    # Colors for each unique demand length
    colors = ['red', 'green', 'blue', 'orange', 'purple', 'cyan', 'brown', 'pink']
    color_map = {}
    color_index = 0

    y_offset = 0

    for i, pattern in enumerate(final_patterns):
        x_start = 0
        y_height = 1

        for piece in pattern:
            # Assign a color for each unique piece length
            if piece not in color_map:
                color_map[piece] = colors[color_index % len(colors)]
                color_index += 1

            rect = patches.Rectangle(
                (x_start, y_offset),     # (x, y)
                piece,                   # width
                y_height,                # height
                facecolor=color_map[piece],
                edgecolor='black'
            )
            ax.add_patch(rect)

            # Label the piece
            ax.text(
                x_start + piece / 2,
                y_offset + 0.4,
                str(piece),
                ha='center',
                va='center',
                fontsize=8,
                color='white'
            )

            x_start += piece

        # Add black rectangle for unused width (waste)
        unused = stock_length - sum(pattern)
        if unused > 0:
            rect = patches.Rectangle(
                (x_start, y_offset),
                unused,
                y_height,
                facecolor='black',
                edgecolor='black'
            )
            ax.add_patch(rect)
            ax.text(
                x_start + unused / 2,
                y_offset + 0.4,
                f"{unused}",
                ha='center',
                va='center',
                fontsize=8,
                color='white'
            )

        y_offset += 2  # Add space before the next pipe

    # Final formatting
    ax.set_xlim(0, stock_length)
    ax.set_ylim(0, y_offset)
    ax.set_xlabel("Length of Pipe")
    ax.set_ylabel("Each Row = One Pipe Used")
    ax.set_title("Cutting Stock Visualization")
    plt.tight_layout()

    # Display this figure in the Tkinter GUI
    # Create a text box to display the results of the optimization
    tk.Label(root, text="Optimization Results:", font=bigFont).grid(row=3, column=0, sticky="w", padx=10, pady=5)
    canvas = FigureCanvasTkAgg(fig, master=root)  # root is the main Tkinter window
    canvas.draw()
    tk.Button(root, text="Save Graph", command=save_graph_image, font=bigFont).grid(row=2, column=1, padx=10, pady=10, sticky="w")
    tk.Button(root, text="Save Results as .txt file", command=save_text_file, font=bigFont).grid(sticky="w", row=2, column=2, padx=10, pady=10)
    canvas.get_tk_widget().grid(row=3, column=0, columnspan=3, padx=10, pady=10)



def save_graph_image():
    if lastFig is None:
        messagebox.showwarning("Warning", "No graph to save yet!")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".png",
                                             filetypes=[("PNG files", "*.png"), ("All Files", "*.*")])
    if file_path:
        try:
            lastFig.savefig(file_path)
            messagebox.showinfo("Success", f"Graph saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save graph: {e}")

def save_text_file():
    global pipes

    standard_length = int(stockPipeLen.get())

    file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[(".txt files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        try:
            with open(file_path, "w") as file:
                file.write("PIPE CUTTING OPTIMIZATION RESULTS\n")
                file.write("=" * 81 + "\n")
                file.write(f"{'Pipe #':<8} {'Cut Length(s)':<50} {'Used':>10} {'Waste':>10}\n")
                file.write("-" * 81 + "\n")

                for i, pipe in enumerate(pipes, 1):
                    used = sum(pipe)
                    waste = standard_length - used
                    cut_lengths_str = ", ".join([str(x) for x in pipe])
                    file.write(f"{i:<8} {cut_lengths_str:<50} {used:>10.1f} {waste:>10.1f}\n")

                file.write("=" * 80 + "\n")
                file.write(f"Total Pipes Used: {len(pipes)}\n")
            messagebox.showinfo("Success", f"Graph saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save graph: {e}")



# Create the main application window using Tkinter
root = tk.Tk()
root.title("Pipe Cutting Optimizer")  # Set the window title
root.geometry("1025x550")



# Create and place the input fields and labels on the window
tk.Label(root, text="Enter Pipe Stock Length (Inches Only)", font = bigFont).grid(row=0, column=0, sticky="w", padx=10, pady=5)
stockPipeLen = tk.StringVar()  # Create a Tkinter StringVar for the stock pipes input
tk.Entry(root, textvariable=stockPipeLen, width=10, font=bigFont).grid(row=0, column=1, padx=10, pady=5)
tk.Label(root, text="Made by Himnish Kaila for Metrix Fabrication", font = ("arial", 12)).grid(row=0, column=2, sticky="w", padx=10, pady=5)

# Create buttons for loading files, optimizing cuts, and saving results
tk.Button(root, text="Load File and Optimize", command=load_optimize, font=bigFont).grid(row=2, column=0, padx=10, pady=10, sticky="w")

# Start the Tkinter main loop to display the GUI window
root.mainloop()
