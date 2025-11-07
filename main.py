import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Global variables for storing cutting patterns and UI font style
pipes = []
bigFont = ("Arial", 20)
lastFig = None  # Holds the latest generated matplotlib figure


def load_optimize():
    """Load an Excel file, perform cutting optimization, and visualize results."""
    # Ask user to select an Excel file
    fileName = filedialog.askopenfilename(
        initialdir="C:",
        title="Open File",
        filetypes=(("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*"))
    )

    # Continue only if a file is selected
    if fileName:
        try:
            # Read Excel file into a pandas DataFrame
            data = pd.read_excel(fileName)

            # Extract quantity and cut length columns
            quantity = data["Qty"].tolist()
            cutLength = data["Size"].tolist()

            # Combine into list of [quantity, length] pairs
            demands = [[quantity[i], cutLength[i]] for i in range(len(quantity))]

            # Get user input for standard stock pipe length
            standard_length = int(stockPipeLen.get())

            # Expand quantities into a flat list of all individual pieces
            pieces = []
            for qty, length in demands:
                for _ in range(qty):
                    pieces.append(length)

            # Sort pieces in descending order to optimize placement
            pieces.sort(reverse=True)

            # Initialize or reset global pipe list
            global pipes
            pipes = []

            # Allocate pieces to pipes using a greedy approach
            for piece in pieces:
                placed = False
                # Try to place in an existing pipe if space allows
                for pipe in pipes:
                    if sum(pipe) + piece <= standard_length:
                        pipe.append(piece)
                        placed = True
                        break
                # Start a new pipe if the piece doesn’t fit anywhere
                if not placed:
                    pipes.append([piece])

            # Draw the final optimized cutting layout
            drawGraph(pipes, standard_length)

        except Exception:
            # Display error if file cannot be read or processed
            messagebox.showerror("Error", "An error occurred while loading the file.")


def drawGraph(final_patterns, stock_length):
    """Display the optimized cutting layout as a bar graph inside the Tkinter GUI."""
    global lastFig
    fig, ax = plt.subplots(figsize=(10, len(final_patterns)))
    lastFig = fig  # Store figure globally for saving later

    # Predefine color palette for visualization
    colors = ['red', 'green', 'blue', 'orange', 'purple', 'cyan', 'brown', 'pink']
    color_map = {}
    color_index = 0
    y_offset = 0  # Used to separate rows vertically for each pipe

    # Draw each pipe as a horizontal bar composed of color-coded segments
    for pattern in final_patterns:
        x_start = 0
        y_height = 1

        for piece in pattern:
            # Assign a consistent color to each unique piece size
            if piece not in color_map:
                color_map[piece] = colors[color_index % len(colors)]
                color_index += 1

            # Create a rectangle representing a single cut piece
            rect = patches.Rectangle(
                (x_start, y_offset), piece, y_height,
                facecolor=color_map[piece], edgecolor='black'
            )
            ax.add_patch(rect)

            # Label each piece with its length
            ax.text(
                x_start + piece / 2, y_offset + 0.4, str(piece),
                ha='center', va='center', fontsize=8, color='white'
            )

            # Update x position for the next piece
            x_start += piece

        # Add black section to represent unused material (waste)
        unused = stock_length - sum(pattern)
        if unused > 0:
            rect = patches.Rectangle(
                (x_start, y_offset), unused, y_height,
                facecolor='black', edgecolor='black'
            )
            ax.add_patch(rect)
            ax.text(
                x_start + unused / 2, y_offset + 0.4, f"{unused}",
                ha='center', va='center', fontsize=8, color='white'
            )

        # Move down to the next row for the next pipe
        y_offset += 2

    # Set chart properties and labels
    ax.set_xlim(0, stock_length)
    ax.set_ylim(0, y_offset)
    ax.set_xlabel("Length of Pipe")
    ax.set_ylabel("Each Row = One Pipe Used")
    ax.set_title("Cutting Stock Visualization")
    plt.tight_layout()

    # Display figure in Tkinter interface
    tk.Label(root, text="Optimization Results:", font=bigFont).grid(row=3, column=0, sticky="w", padx=10, pady=5)
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()

    # Add buttons for saving graph and text output
    tk.Button(root, text="Save Graph", command=save_graph_image, font=bigFont).grid(row=2, column=1, padx=10, pady=10, sticky="w")
    tk.Button(root, text="Save Results as .txt file", command=save_text_file, font=bigFont).grid(row=2, column=2, padx=10, pady=10, sticky="w")
    canvas.get_tk_widget().grid(row=3, column=0, columnspan=3, padx=10, pady=10)


def save_graph_image():
    """Save the currently displayed matplotlib graph as a PNG file."""
    # Ensure a figure exists before attempting to save
    if lastFig is None:
        messagebox.showwarning("Warning", "No graph to save yet!")
        return

    # Prompt user for save location
    file_path = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("PNG files", "*.png"), ("All Files", "*.*")]
    )

    # Save the figure if a valid path is selected
    if file_path:
        try:
            lastFig.savefig(file_path)
            messagebox.showinfo("Success", f"Graph saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save graph: {e}")


def save_text_file():
    """Save the cutting results and waste calculations as a formatted text file."""
    global pipes
    standard_length = int(stockPipeLen.get())

    # Prompt user for output filename
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All Files", "*.*")]
    )

    # Generate report only if valid path provided
    if file_path:
        try:
            with open(file_path, "w") as file:
                file.write("PIPE CUTTING OPTIMIZATION RESULTS\n")
                file.write("=" * 81 + "\n")
                file.write(f"{'Pipe #':<8} {'Cut Length(s)':<50} {'Used':>10} {'Waste':>10}\n")
                file.write("-" * 81 + "\n")

                # Write details for each optimized pipe
                for i, pipe in enumerate(pipes, 1):
                    used = sum(pipe)
                    waste = standard_length - used
                    cut_lengths_str = ", ".join([str(x) for x in pipe])
                    file.write(f"{i:<8} {cut_lengths_str:<50} {used:>10.1f} {waste:>10.1f}\n")

                # Summary line
                file.write("=" * 80 + "\n")
                file.write(f"Total Pipes Used: {len(pipes)}\n")

            messagebox.showinfo("Success", f"Results saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save text file: {e}")


# ---------------------------
# Tkinter GUI Setup
# ---------------------------

# Create and configure main window
root = tk.Tk()
root.title("Pipe Cutting Optimizer")
root.geometry("1025x550")

# Label and entry for standard pipe length input
tk.Label(root, text="Enter Pipe Stock Length (Inches Only)", font=bigFont).grid(row=0, column=0, sticky="w", padx=10, pady=5)
stockPipeLen = tk.StringVar()
tk.Entry(root, textvariable=stockPipeLen, width=10, font=bigFont).grid(row=0, column=1, padx=10, pady=5)

# Developer credit label
tk.Label(root, text="REDACTED", font=("Arial", 12)).grid(row=0, column=2, sticky="w", padx=10, pady=5)

# Button to load file and begin optimization
tk.Button(root, text="Load File and Optimize", command=load_optimize, font=bigFont).grid(row=2, column=0, padx=10, pady=10, sticky="w")

# Start the GUI event loop
root.mainloop()
