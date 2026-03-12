#!/usr/bin/env python3
"""
Real-time Windows screen capture using PowerShell + Python
"""
import subprocess
import time
import os
from PIL import Image
import io

class WindowsScreenCapture:
    def __init__(self, output_dir='/tmp/screen_stream'):
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        self.frame_count = 0
        
    def capture_frame(self):
        """Capture single frame using PowerShell"""
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        filename = f'{self.output_dir}/frame_{timestamp}_{self.frame_count:04d}.png'
        
        # PowerShell command to capture screen
        ps_cmd = f'''
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bitmap.Save('{filename}')
$graphics.Dispose()
$bitmap.Dispose()
'''
        
        # Run PowerShell
        result = subprocess.run(
            ['/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe', 
             '-Command', ps_cmd],
            capture_output=True
        )
        
        if result.returncode == 0 and os.path.exists(filename):
            self.frame_count += 1
            return filename
        else:
            print(f"Capture failed: {result.stderr}")
            return None
    
    def start_stream(self, fps=5, duration=None):
        """Start capturing stream"""
        print(f"Starting screen stream ({fps} FPS)...")
        print(f"Output: {self.output_dir}")
        print("Press Ctrl+C to stop")
        
        frame_time = 1.0 / fps
        start_time = time.time()
        
        try:
            while True:
                loop_start = time.time()
                
                # Capture frame
                filename = self.capture_frame()
                if filename:
                    print(f"Captured: {os.path.basename(filename)}")
                
                # Check duration
                if duration and (time.time() - start_time) >= duration:
                    print(f"Reached duration limit: {duration}s")
                    break
                
                # Maintain FPS
                elapsed = time.time() - loop_start
                if elapsed < frame_time:
                    time.sleep(frame_time - elapsed)
                    
        except KeyboardInterrupt:
            print("\nStream stopped")
            
    def get_latest_frame(self):
        """Get the most recent frame"""
        files = sorted([f for f in os.listdir(self.output_dir) if f.endswith('.png')])
        if files:
            return os.path.join(self.output_dir, files[-1])
        return None


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Windows Screen Capture')
    parser.add_argument('--fps', type=int, default=5, help='Frames per second')
    parser.add_argument('--duration', type=int, help='Duration in seconds')
    parser.add_argument('--output', default='/tmp/screen_stream', help='Output directory')
    args = parser.parse_args()
    
    capture = WindowsScreenCapture(args.output)
    capture.start_stream(fps=args.fps, duration=args.duration)


if __name__ == '__main__':
    main()
