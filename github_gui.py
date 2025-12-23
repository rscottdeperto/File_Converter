import tkinter as tk
from tkinter import messagebox
import subprocess


class GitHubGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple GitHub GUI")
        self.root.geometry("400x250")

        tk.Label(root, text="Commit Message:").pack(pady=5)
        self.commit_entry = tk.Entry(root, width=50)
        self.commit_entry.pack(pady=5)

        self.status_text = tk.Text(root, height=7, width=50, state='disabled')
        self.status_text.pack(pady=5)

        tk.Button(root, text="Add All Files",
                  command=self.git_add).pack(pady=2)
        tk.Button(root, text="Commit", command=self.git_commit).pack(pady=2)
        tk.Button(root, text="Push", command=self.git_push).pack(pady=2)
        tk.Button(root, text="Init Git LFS & Track Large Files",
                  command=self.git_lfs_setup).pack(pady=2)

    def git_lfs_setup(self):
        # List of common large file types to track
        large_types = ["*.zip", "*.csv", "*.exe", "*.xls", "*.xlsx"]
        # Initialize Git LFS
        success = self.run_git_command(["git", "lfs", "install"])
        if not success:
            return
        for ext in large_types:
            self.run_git_command(["git", "lfs", "track", ext])
        self.status_text.config(state='normal')
        self.status_text.insert(
            tk.END, "Git LFS initialized and large file types tracked.\n")
        self.status_text.config(state='disabled')
        self.status_text.see(tk.END)

    def run_git_command(self, command):
        try:
            result = subprocess.run(
                command, capture_output=True, text=True, shell=True)
            output = result.stdout + result.stderr
            self.status_text.config(state='normal')
            self.status_text.insert(
                tk.END, f"$ {' '.join(command)}\n{output}\n")
            self.status_text.config(state='disabled')
            self.status_text.see(tk.END)
            return result.returncode == 0
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return False

    def git_add(self):
        self.run_git_command(["git", "add", "."])

    def git_commit(self):
        msg = self.commit_entry.get()
        if not msg:
            messagebox.showwarning("Warning", "Please enter a commit message.")
            return
        self.run_git_command(["git", "commit", "-m", msg])

    def git_push(self):
        self.run_git_command(["git", "push"])


if __name__ == "__main__":
    root = tk.Tk()
    app = GitHubGUI(root)
    root.mainloop()
