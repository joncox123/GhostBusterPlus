using System;
using System.Diagnostics;
using SharpDX;
using SharpDX.DXGI;
using SharpDX.Direct3D11;
using SharpDX.Direct3D;
using SharpDX.D3DCompiler;

namespace ScreenRefreshApp
{
    /// <summary>
    /// Handles DirectX-based screenshot processing, including capture, grayscale conversion, and comparison.
    /// This class uses GPU acceleration via DirectX 11 compute shaders to efficiently process screen content.
    /// </summary>
    public class Processor : System.IDisposable
    {
        // Core DirectX device objects
        private SharpDX.Direct3D11.Device d3dDevice;                // The main Direct3D device for creating resources and executing commands
        private SharpDX.DXGI.OutputDuplication outputDuplication;   // Screen duplication interface for capturing desktop content
        private SharpDX.DXGI.Factory1 dxgiFactory;                  // Factory for creating DXGI objects
        private SharpDX.DXGI.Adapter adapter;                       // Graphics adapter (GPU)
        private SharpDX.DXGI.Output output;                         // Monitor output

        // Compute shaders for GPU-based processing
        private SharpDX.Direct3D11.ComputeShader grayscaleComputeShader; // Shader that converts color images to grayscale
        private SharpDX.Direct3D11.ComputeShader compareComputeShader;   // Shader that compares two grayscale images

        // Resource views for GPU memory access
        private SharpDX.Direct3D11.UnorderedAccessView grayscaleUAV;     // Allows compute shader to write to grayscale texture
        private SharpDX.Direct3D11.ShaderResourceView previousGrayscaleSRV;  // Previous frame grayscale data for reading in shader
        private SharpDX.Direct3D11.ShaderResourceView currentGrayscaleSRV;   // Current frame grayscale data for reading in shader
        private SharpDX.Direct3D11.UnorderedAccessView resultUAV;        // For writing comparison results

        // Buffers for data storage and transfer
        private SharpDX.Direct3D11.Buffer resultBuffer;             // GPU-side buffer for storing comparison results
        private SharpDX.Direct3D11.Buffer resultStagingBuffer;      // CPU-readable buffer for transferring results from GPU

        // Texture resources for image data
        private SharpDX.Direct3D11.Texture2D previousGrayscaleTexture;  // Stores previous frame in grayscale
        private SharpDX.Direct3D11.Texture2D currentGrayscaleTexture;   // Stores current frame in grayscale

        // Configuration
        private double pixelThresholdPct = 3.0;  // Default threshold for significant change (percentage of pixels)

        /// <summary>
        /// Gets or sets the percentage of pixels that must change for a refresh to be triggered.
        /// Higher values mean more change is required to trigger a refresh.
        /// </summary>
        public double PixelThresholdPct
        {
            get { return pixelThresholdPct; }
            set { pixelThresholdPct = value; }
        }

        /// <summary>
        /// Initializes a new instance of the Processor class and sets up DirectX resources.
        /// </summary>
        public Processor()
        {
            InitializeDirectX();
        }

        /// <summary>
        /// Initializes DirectX resources for screenshot processing.
        /// This includes setting up the GPU device, creating textures, buffers, and compiling shaders.
        /// </summary>
        private void InitializeDirectX()
        {
            try
            {
                // Create a DXGI factory - the entry point for DirectX Graphics Infrastructure
                dxgiFactory = new SharpDX.DXGI.Factory1();

                // Get the primary graphics adapter (typically the dedicated GPU)
                adapter = dxgiFactory.GetAdapter1(0);

                // Create a DirectX device with the debug layer enabled to help catch errors
                // BgraSupport is needed for desktop duplication which uses BGRA format
                d3dDevice = new SharpDX.Direct3D11.Device(adapter,
                    SharpDX.Direct3D11.DeviceCreationFlags.Debug |
                    SharpDX.Direct3D11.DeviceCreationFlags.BgraSupport);
                System.Console.WriteLine("DirectX device created with debug layer enabled.");
                System.Console.WriteLine($"Device feature level: {d3dDevice.FeatureLevel}");

                // Get the primary monitor output and set up desktop duplication
                output = adapter.GetOutput(0);
                var output1 = output.QueryInterface<SharpDX.DXGI.Output1>();
                outputDuplication = output1.DuplicateOutput(d3dDevice);
                System.Console.WriteLine("Output duplication initialized.");

                // Get the desktop dimensions to create appropriately sized textures
                int width = output.Description.DesktopBounds.Right - output.Description.DesktopBounds.Left;
                int height = output.Description.DesktopBounds.Bottom - output.Description.DesktopBounds.Top;
                System.Console.WriteLine($"Texture dimensions: Width={width}, Height={height}");

                // Create texture description for grayscale images
                // Using R32_Float format for single-channel grayscale values (32-bit float per pixel)
                var grayscaleDesc = new SharpDX.Direct3D11.Texture2DDescription
                {
                    Width = width,
                    Height = height,
                    MipLevels = 1,
                    ArraySize = 1,
                    Format = SharpDX.DXGI.Format.R32_Float,             // Single-channel 32-bit float format
                    SampleDescription = new SharpDX.DXGI.SampleDescription(1, 0), // No multisampling
                    Usage = SharpDX.Direct3D11.ResourceUsage.Default,   // GPU-readable and writable
                    BindFlags = SharpDX.Direct3D11.BindFlags.UnorderedAccess | // Allow compute shader writes
                                SharpDX.Direct3D11.BindFlags.ShaderResource,   // Allow shader reads
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.None,  // No CPU access needed
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.None
                };

                // Create two textures - one for current frame, one for previous frame
                currentGrayscaleTexture = new SharpDX.Direct3D11.Texture2D(d3dDevice, grayscaleDesc);
                previousGrayscaleTexture = new SharpDX.Direct3D11.Texture2D(d3dDevice, grayscaleDesc);

                // Create shader resource views (SRVs) for reading the grayscale textures in shaders
                currentGrayscaleSRV = new SharpDX.Direct3D11.ShaderResourceView(d3dDevice,
                    currentGrayscaleTexture,
                    new SharpDX.Direct3D11.ShaderResourceViewDescription
                    {
                        Format = SharpDX.DXGI.Format.R32_Float,
                        Dimension = SharpDX.Direct3D.ShaderResourceViewDimension.Texture2D,
                        Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
                    });

                previousGrayscaleSRV = new SharpDX.Direct3D11.ShaderResourceView(d3dDevice,
                    previousGrayscaleTexture,
                    new SharpDX.Direct3D11.ShaderResourceViewDescription
                    {
                        Format = SharpDX.DXGI.Format.R32_Float,
                        Dimension = SharpDX.Direct3D.ShaderResourceViewDimension.Texture2D,
                        Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
                    });

                // Create unordered access view (UAV) for writing to the current grayscale texture
                grayscaleUAV = new SharpDX.Direct3D11.UnorderedAccessView(d3dDevice,
                    currentGrayscaleTexture,
                    new SharpDX.Direct3D11.UnorderedAccessViewDescription
                    {
                        Format = SharpDX.DXGI.Format.R32_Float,
                        Dimension = SharpDX.Direct3D11.UnorderedAccessViewDimension.Texture2D,
                        Texture2D = { MipSlice = 0 }
                    });
                System.Console.WriteLine("Grayscale textures and views created.");

                // Create buffer for storing pixel difference count in the comparison shader
                // The result buffer is a structured buffer that will contain one integer count
                var resultDesc = new SharpDX.Direct3D11.BufferDescription
                {
                    Usage = SharpDX.Direct3D11.ResourceUsage.Default,     // GPU-readable and writable
                    SizeInBytes = sizeof(int),                             // Just need one int for the diff count
                    BindFlags = SharpDX.Direct3D11.BindFlags.UnorderedAccess, // For compute shader writes
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.None,  // No direct CPU access
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.BufferStructured, // Structured buffer
                    StructureByteStride = sizeof(int)                      // Each element is one int
                };
                resultBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, resultDesc);

                // Create staging buffer for reading back the result from GPU to CPU
                // Staging resources allow efficient transfer from GPU to CPU memory
                var resultStagingDesc = new SharpDX.Direct3D11.BufferDescription
                {
                    Usage = SharpDX.Direct3D11.ResourceUsage.Staging,      // For CPU read-back
                    SizeInBytes = sizeof(int),
                    BindFlags = SharpDX.Direct3D11.BindFlags.None,         // No shader bindings needed
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.Read, // Allow CPU reads
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.BufferStructured,
                    StructureByteStride = sizeof(int)
                };
                resultStagingBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, resultStagingDesc);
                resultUAV = new SharpDX.Direct3D11.UnorderedAccessView(d3dDevice, resultBuffer);
                System.Console.WriteLine("Result buffers and UAV created.");

                // Define HLSL shader for grayscale conversion
                // This shader takes a color texture (RGBA) and outputs a single-channel grayscale texture
                string grayscaleShaderCode = @"
                    // Input color texture (RGBA format)
                    Texture2D<float4> InputTexture : register(t0);
                    
                    // Output grayscale texture (single channel)
                    RWTexture2D<float> OutputTexture : register(u0);
                    
                    // Constant buffer with texture dimensions
                    cbuffer TextureSizeBuffer : register(b0)
                    {
                        uint Width;   // Texture width in pixels
                        uint Height;  // Texture height in pixels
                    }
                    
                    // Compute shader entry point
                    // numthreads defines the thread group size (16x16 threads per group)
                    [numthreads(16, 16, 1)]
                    void main(uint3 dispatchThreadID : SV_DispatchThreadID)
                    {
                        // Get the pixel coordinates from the thread ID
                        uint2 texCoord = dispatchThreadID.xy;
                        
                        // Only process pixels within the texture bounds
                        if (texCoord.x < Width && texCoord.y < Height)
                        {
                            // Sample the color from the input texture
                            float4 color = InputTexture[texCoord];
                            
                            // Convert to grayscale using a simple average of RGB components
                            // (More precise conversion could use weighted coefficients)
                            float gray = (color.r + color.g + color.b) / 3.0;
                            
                            // Write the grayscale value to the output texture
                            OutputTexture[texCoord] = gray;
                        }
                    }";

                // Define HLSL shader for comparing two grayscale images
                // This shader counts pixels that have changed by more than a threshold
                string compareShaderCode = @"
                    // Previous frame grayscale texture
                    Texture2D<float> OldTexture : register(t0);
                    
                    // Current frame grayscale texture
                    Texture2D<float> NewTexture : register(t1);
                    
                    // Output buffer for counting changed pixels
                    RWStructuredBuffer<uint> ResultBuffer : register(u0);
                    
                    // Constant buffer with texture dimensions
                    cbuffer TextureSizeBuffer : register(b0)
                    {
                        uint Width;   // Texture width in pixels
                        uint Height;  // Texture height in pixels
                    }
                    
                    // Compute shader entry point
                    [numthreads(16, 16, 1)]
                    void main(uint3 dispatchThreadID : SV_DispatchThreadID)
                    {
                        // Get the pixel coordinates from the thread ID
                        uint2 texCoord = dispatchThreadID.xy;
                        
                        // Only process pixels within the texture bounds
                        if (texCoord.x < Width && texCoord.y < Height)
                        {
                            // Get the grayscale values from both textures
                            float oldPixel = OldTexture[texCoord];
                            float newPixel = NewTexture[texCoord];
                            
                            // Calculate the absolute difference
                            float diff = abs(oldPixel - newPixel);
                            
                            // If the difference exceeds our threshold (0.05 ≈ 5% brightness change)
                            // count this pixel as changed
                            if (diff > 0.05)
                            {
                                // Atomically increment the counter in the result buffer
                                // InterlockedAdd ensures correct counting with multiple threads
                                InterlockedAdd(ResultBuffer[0], 1);
                            }
                        }
                    }";

                // Compile the HLSL shaders into bytecode that the GPU can execute
                using (var grayscaleBytecode = SharpDX.D3DCompiler.ShaderBytecode.Compile(grayscaleShaderCode, "main", "cs_5_0"))
                using (var compareBytecode = SharpDX.D3DCompiler.ShaderBytecode.Compile(compareShaderCode, "main", "cs_5_0"))
                {
                    // Check if compilation was successful
                    if (grayscaleBytecode.Bytecode != null && compareBytecode.Bytecode != null)
                    {
                        // Create compute shader objects from the compiled bytecode
                        grayscaleComputeShader = new SharpDX.Direct3D11.ComputeShader(d3dDevice, grayscaleBytecode.Bytecode);
                        compareComputeShader = new SharpDX.Direct3D11.ComputeShader(d3dDevice, compareBytecode.Bytecode);
                        System.Console.WriteLine("Compute shaders compiled successfully.");
                    }
                    else
                    {
                        // Report compilation errors
                        string errorMsg = "Shader compilation failed: ";
                        errorMsg += grayscaleBytecode.Bytecode == null ? $"Grayscale shader error: {grayscaleBytecode.Message}" : "";
                        errorMsg += compareBytecode.Bytecode == null ? $"Compare shader error: {compareBytecode.Message}" : "";
                        throw new System.Exception(errorMsg);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine($"DirectX initialization failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Captures a screenshot, converts it to grayscale, and compares it with the previous frame.
        /// Returns true if the difference exceeds the threshold percentage.
        /// </summary>
        /// <returns>True if significant screen change is detected, otherwise false.</returns>
        public bool ProcessScreenshotOnGPU()
        {
            //System.Console.WriteLine($"Screenshot taken at {System.DateTime.Now:HH:mm:ss.fff}");
            SharpDX.DXGI.Resource desktopResource = null;
            SharpDX.DXGI.OutputDuplicateFrameInformation frameInfo;
            bool frameAcquired = false;

            try
            {
                // Attempt to acquire the next desktop frame with retries
                // Frame acquisition can fail if there have been no updates or if the desktop is locked
                int retryCount = 0;
                const int maxRetries = 3;
                while (retryCount < maxRetries && !frameAcquired)
                {
                    try
                    {
                        // AcquireNextFrame waits for up to 500ms for a new frame
                        outputDuplication.AcquireNextFrame(500, out frameInfo, out desktopResource);
                        frameAcquired = true;
                    }
                    catch (SharpDX.SharpDXException ex) when (ex.HResult == unchecked((int)0x887A0001))
                    {
                        // 0x887A0001 = DXGI_ERROR_WAIT_TIMEOUT - No new frame available yet
                        System.Console.WriteLine($"AcquireNextFrame failed (attempt {retryCount + 1}/{maxRetries}): {ex.Message}");
                        retryCount++;
                        if (retryCount == maxRetries)
                        {
                            throw new System.Exception("Failed to acquire frame after retries", ex);
                        }
                        System.Threading.Thread.Sleep(100);
                    }
                }

                if (!frameAcquired)
                {
                    throw new System.Exception("Unable to acquire frame after retries.");
                }

                // Convert the desktop resource to a texture we can work with
                using (var capturedTexture = desktopResource.QueryInterface<SharpDX.Direct3D11.Texture2D>())
                {
                    // Verify dimensions match our output textures
                    if (capturedTexture.Description.Width != currentGrayscaleTexture.Description.Width ||
                        capturedTexture.Description.Height != currentGrayscaleTexture.Description.Height)
                    {
                        System.Console.WriteLine($"Captured texture dimensions mismatch: Expected {currentGrayscaleTexture.Description.Width}x{currentGrayscaleTexture.Description.Height}, Got {capturedTexture.Description.Width}x{capturedTexture.Description.Height}");
                        return false;
                    }

                    // Convert the color desktop image to grayscale using our compute shader
                    ConvertToGrayscaleGPU(capturedTexture);

                    bool significantChange = false;
                    if (previousGrayscaleTexture != null)
                    {
                        // Compare the current grayscale image with the previous one
                        int diffCount = CompareScreenshotsGPU();
                        int totalPixels = currentGrayscaleTexture.Description.Width * currentGrayscaleTexture.Description.Height;
                        double diffPercentage = (double)diffCount / totalPixels * 100.0;
                        //System.Console.WriteLine($"CompareScreenshotsGPU ran at {System.DateTime.Now:HH:mm:ss.fff}, DiffCount: {diffCount}, Total Pixels: {totalPixels}, Percentage: {diffPercentage:F2}%");

                        // Compare against the configurable threshold
                        significantChange = diffPercentage >= pixelThresholdPct;
                    }
                    else
                    {
                        System.Console.WriteLine("Previous grayscale texture not available, skipping comparison.");
                    }

                    // Save the current frame as the previous frame for the next comparison
                    d3dDevice.ImmediateContext.CopyResource(currentGrayscaleTexture, previousGrayscaleTexture);

                    return significantChange;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Screenshot processing failed: {ex.Message}");
                System.Console.WriteLine($"Screenshot processing failed: {ex.Message}");
                return false;
            }
            finally
            {
                // Always release the frame and dispose resources, even if an exception occurred
                if (frameAcquired)
                {
                    try
                    {
                        outputDuplication.ReleaseFrame();
                    }
                    catch (System.Exception ex)
                    {
                        System.Console.WriteLine($"Failed to release frame: {ex.Message}");
                    }
                }
                else
                {
                    System.Console.WriteLine("Frame not acquired, skipping release.");
                }
                desktopResource?.Dispose();
            }
        }

        /// <summary>
        /// Converts a color texture to grayscale using a GPU compute shader.
        /// This is much faster than CPU-based conversion for large images.
        /// </summary>
        /// <param name="inputTexture">The color texture to convert (typically B8G8R8A8_UNorm format)</param>
        private void ConvertToGrayscaleGPU(SharpDX.Direct3D11.Texture2D inputTexture)
        {
            // Verify the input texture is in the expected format (BGRA)
            if (inputTexture.Description.Format != SharpDX.DXGI.Format.B8G8R8A8_UNorm)
            {
                System.Console.WriteLine($"Invalid input texture format: Expected B8G8R8A8_UNorm, Got {inputTexture.Description.Format}");
                throw new System.InvalidOperationException("Input texture format must be B8G8R8A8_UNorm.");
            }

            var context = d3dDevice.ImmediateContext;

            // Unbind any previous resources to avoid conflicts
            // DirectX requires explicit unbinding of resources between different operations
            context.ComputeShader.SetUnorderedAccessView(0, null);
            context.ComputeShader.SetShaderResource(0, null);
            context.ComputeShader.SetConstantBuffer(0, null);

            // Create a shader resource view for the input texture
            // This view lets the shader read from the desktop texture
            var inputSRV = new ShaderResourceView(d3dDevice, inputTexture, new ShaderResourceViewDescription
            {
                Format = Format.B8G8R8A8_UNorm,        // BGRA 8-bit normalized format
                Dimension = ShaderResourceViewDimension.Texture2D,
                Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
            });

            SharpDX.Direct3D11.Buffer textureSizeBuffer = null;
            try
            {
                // Create a constant buffer to pass texture dimensions to the shader
                // Constant buffers must be a multiple of 16 bytes in size (DirectX requirement)
                textureSizeBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, new BufferDescription
                {
                    SizeInBytes = 16,                  // 16 bytes = 4 uint values (minimum size)
                    Usage = ResourceUsage.Default,     // GPU read-only
                    BindFlags = BindFlags.ConstantBuffer,
                    CpuAccessFlags = CpuAccessFlags.None
                });

                // Prepare the data for the constant buffer: width, height, and padding
                var textureSizeData = new[] {
                    (uint)inputTexture.Description.Width,
                    (uint)inputTexture.Description.Height,
                    0u, 0u  // Padding to 16 bytes (DirectX requirement)
                };

                // Update the constant buffer with the texture dimensions
                context.UpdateSubresource(textureSizeData, textureSizeBuffer);

                // Set the compute shader - this specifies which program will run on the GPU
                context.ComputeShader.Set(grayscaleComputeShader);

                // Bind the constant buffer containing texture dimensions
                context.ComputeShader.SetConstantBuffer(0, textureSizeBuffer);

                // Bind the input texture for the shader to read from
                context.ComputeShader.SetShaderResource(0, inputSRV);

                // Bind the output texture for the shader to write to
                context.ComputeShader.SetUnorderedAccessView(0, grayscaleUAV);

                // Calculate how many thread groups to dispatch
                // Each thread group processes 16x16 pixels (as defined in the shader)
                // We round up to ensure all pixels are covered
                int threadGroupX = (inputTexture.Description.Width + 15) / 16;
                int threadGroupY = (inputTexture.Description.Height + 15) / 16;

                // Launch the compute shader on the GPU
                context.Dispatch(threadGroupX, threadGroupY, 1);

                // Unbind resources after use to avoid conflicts with subsequent operations
                context.ComputeShader.SetConstantBuffer(0, null);
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
            }
            catch (SharpDX.SharpDXException ex)
            {
                System.Console.WriteLine($"Compute shader setup failed: {ex.Message}");
                throw;
            }
            finally
            {
                // Always dispose temporary resources to prevent memory leaks
                inputSRV.Dispose();
                textureSizeBuffer?.Dispose();
            }
        }

        /// <summary>
        /// Compares the previous and current grayscale textures using a GPU compute shader.
        /// Counts the number of pixels that differ by more than a threshold.
        /// </summary>
        /// <returns>Number of pixels that changed significantly between frames</returns>
        private int CompareScreenshotsGPU()
        {
            var context = d3dDevice.ImmediateContext;

            // Reset the result buffer to zero
            int zero = 0;
            context.UpdateSubresource(ref zero, resultBuffer);

            // Create a constant buffer for texture dimensions
            // This needs to be a multiple of 16 bytes (DirectX requirement)
            using (var textureSizeBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, new BufferDescription
            {
                SizeInBytes = 16,                  // 16 bytes = 4 uint values
                Usage = ResourceUsage.Default,     // GPU read-only
                BindFlags = BindFlags.ConstantBuffer,
                CpuAccessFlags = CpuAccessFlags.None
            }))
            {
                // Prepare the data for the constant buffer: width, height, and padding
                var textureSizeData = new[] {
                    (uint)currentGrayscaleTexture.Description.Width,
                    (uint)currentGrayscaleTexture.Description.Height,
                    0u, 0u  // Padding to 16 bytes (DirectX requirement)
                };

                // Update the constant buffer with the texture dimensions
                context.UpdateSubresource(textureSizeData, textureSizeBuffer);

                // Unbind any previous resources to avoid conflicts
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetShaderResource(1, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
                context.ComputeShader.SetConstantBuffer(0, null);

                // Set the comparison compute shader
                context.ComputeShader.Set(compareComputeShader);

                // Bind the constant buffer containing texture dimensions
                context.ComputeShader.SetConstantBuffer(0, textureSizeBuffer);

                // Bind the previous and current textures for comparison
                context.ComputeShader.SetShaderResource(0, previousGrayscaleSRV);
                context.ComputeShader.SetShaderResource(1, currentGrayscaleSRV);

                // Bind the result buffer for the shader to write the diff count
                context.ComputeShader.SetUnorderedAccessView(0, resultUAV);

                // Calculate how many thread groups to dispatch
                // Each thread group processes 16x16 pixels (as defined in the shader)
                int threadGroupX = (currentGrayscaleTexture.Description.Width + 15) / 16;
                int threadGroupY = (currentGrayscaleTexture.Description.Height + 15) / 16;

                // Launch the compute shader on the GPU
                context.Dispatch(threadGroupX, threadGroupY, 1);

                // Unbind resources after use
                context.ComputeShader.SetConstantBuffer(0, null);
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetShaderResource(1, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
            }

            // Copy the result from the GPU buffer to a CPU-readable buffer
            context.CopyResource(resultBuffer, resultStagingBuffer);

            // Map the staging buffer to get CPU access to the data
            var dataBox = context.MapSubresource(resultStagingBuffer, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);

            // Read the diff count from the buffer
            int diffCount = System.Runtime.InteropServices.Marshal.ReadInt32(dataBox.DataPointer);

            // Unmap to release the resource
            context.UnmapSubresource(resultStagingBuffer, 0);

            return diffCount;
        }

        /// <summary>
        /// Disposes of all DirectX resources.
        /// This must be called when the processor is no longer needed to avoid memory leaks.
        /// </summary>
        public void Dispose()
        {
            // Dispose textures
            previousGrayscaleTexture?.Dispose();
            currentGrayscaleTexture?.Dispose();

            // Dispose views
            grayscaleUAV?.Dispose();
            previousGrayscaleSRV?.Dispose();
            currentGrayscaleSRV?.Dispose();
            resultUAV?.Dispose();

            // Dispose buffers
            resultBuffer?.Dispose();
            resultStagingBuffer?.Dispose();

            // Dispose shaders
            grayscaleComputeShader?.Dispose();
            compareComputeShader?.Dispose();

            // Dispose DXGI resources
            outputDuplication?.Dispose();
            output?.Dispose();
            adapter?.Dispose();
            dxgiFactory?.Dispose();

            // Dispose the device last
            d3dDevice?.Dispose();
        }
    }
}