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
    /// </summary>
    public class Processor : System.IDisposable
    {
        private SharpDX.Direct3D11.Device d3dDevice;
        private SharpDX.DXGI.OutputDuplication outputDuplication;
        private SharpDX.DXGI.Factory1 dxgiFactory;
        private SharpDX.DXGI.Adapter adapter;
        private SharpDX.DXGI.Output output;
        private SharpDX.Direct3D11.ComputeShader grayscaleComputeShader;
        private SharpDX.Direct3D11.ComputeShader compareComputeShader;
        private SharpDX.Direct3D11.UnorderedAccessView grayscaleUAV;
        private SharpDX.Direct3D11.ShaderResourceView previousGrayscaleSRV;
        private SharpDX.Direct3D11.ShaderResourceView currentGrayscaleSRV;
        private SharpDX.Direct3D11.UnorderedAccessView resultUAV;
        private SharpDX.Direct3D11.Buffer resultBuffer;
        private SharpDX.Direct3D11.Buffer resultStagingBuffer;
        private SharpDX.Direct3D11.Texture2D previousGrayscaleTexture;
        private SharpDX.Direct3D11.Texture2D currentGrayscaleTexture;
        private const double N_PIXEL_THRESH_PCT = 5.0; // Threshold for significant change

        public Processor()
        {
            InitializeDirectX();
        }

        /// <summary>
        /// Initializes DirectX resources for screenshot processing.
        /// </summary>
        private void InitializeDirectX()
        {
            try
            {
                dxgiFactory = new SharpDX.DXGI.Factory1();
                adapter = dxgiFactory.GetAdapter1(0);
                d3dDevice = new SharpDX.Direct3D11.Device(adapter, SharpDX.Direct3D11.DeviceCreationFlags.Debug | SharpDX.Direct3D11.DeviceCreationFlags.BgraSupport);
                System.Console.WriteLine("DirectX device created with debug layer enabled.");
                System.Console.WriteLine($"Device feature level: {d3dDevice.FeatureLevel}");

                output = adapter.GetOutput(0);
                var output1 = output.QueryInterface<SharpDX.DXGI.Output1>();
                outputDuplication = output1.DuplicateOutput(d3dDevice);
                System.Console.WriteLine("Output duplication initialized.");

                int width = output.Description.DesktopBounds.Right - output.Description.DesktopBounds.Left;
                int height = output.Description.DesktopBounds.Bottom - output.Description.DesktopBounds.Top;

                System.Console.WriteLine($"Texture dimensions: Width={width}, Height={height}");

                var grayscaleDesc = new SharpDX.Direct3D11.Texture2DDescription
                {
                    Width = width,
                    Height = height,
                    MipLevels = 1,
                    ArraySize = 1,
                    Format = SharpDX.DXGI.Format.R32_Float,
                    SampleDescription = new SharpDX.DXGI.SampleDescription(1, 0),
                    Usage = SharpDX.Direct3D11.ResourceUsage.Default,
                    BindFlags = SharpDX.Direct3D11.BindFlags.UnorderedAccess | SharpDX.Direct3D11.BindFlags.ShaderResource,
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.None,
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.None
                };

                currentGrayscaleTexture = new SharpDX.Direct3D11.Texture2D(d3dDevice, grayscaleDesc);
                previousGrayscaleTexture = new SharpDX.Direct3D11.Texture2D(d3dDevice, grayscaleDesc);

                currentGrayscaleSRV = new SharpDX.Direct3D11.ShaderResourceView(d3dDevice, currentGrayscaleTexture, new SharpDX.Direct3D11.ShaderResourceViewDescription
                {
                    Format = SharpDX.DXGI.Format.R32_Float,
                    Dimension = SharpDX.Direct3D.ShaderResourceViewDimension.Texture2D,
                    Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
                });
                previousGrayscaleSRV = new SharpDX.Direct3D11.ShaderResourceView(d3dDevice, previousGrayscaleTexture, new SharpDX.Direct3D11.ShaderResourceViewDescription
                {
                    Format = SharpDX.DXGI.Format.R32_Float,
                    Dimension = SharpDX.Direct3D.ShaderResourceViewDimension.Texture2D,
                    Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
                });

                grayscaleUAV = new SharpDX.Direct3D11.UnorderedAccessView(d3dDevice, currentGrayscaleTexture, new SharpDX.Direct3D11.UnorderedAccessViewDescription
                {
                    Format = SharpDX.DXGI.Format.R32_Float,
                    Dimension = SharpDX.Direct3D11.UnorderedAccessViewDimension.Texture2D,
                    Texture2D = { MipSlice = 0 }
                });
                System.Console.WriteLine("Grayscale textures and views created.");

                var resultDesc = new SharpDX.Direct3D11.BufferDescription
                {
                    Usage = SharpDX.Direct3D11.ResourceUsage.Default,
                    SizeInBytes = sizeof(int),
                    BindFlags = SharpDX.Direct3D11.BindFlags.UnorderedAccess,
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.None,
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.BufferStructured,
                    StructureByteStride = sizeof(int)
                };
                resultBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, resultDesc);

                var resultStagingDesc = new SharpDX.Direct3D11.BufferDescription
                {
                    Usage = SharpDX.Direct3D11.ResourceUsage.Staging,
                    SizeInBytes = sizeof(int),
                    BindFlags = SharpDX.Direct3D11.BindFlags.None,
                    CpuAccessFlags = SharpDX.Direct3D11.CpuAccessFlags.Read,
                    OptionFlags = SharpDX.Direct3D11.ResourceOptionFlags.BufferStructured,
                    StructureByteStride = sizeof(int)
                };
                resultStagingBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, resultStagingDesc);
                resultUAV = new SharpDX.Direct3D11.UnorderedAccessView(d3dDevice, resultBuffer);
                System.Console.WriteLine("Result buffers and UAV created.");

                string grayscaleShaderCode = @"
                    Texture2D<float4> InputTexture : register(t0);
                    RWTexture2D<float> OutputTexture : register(u0);
                    cbuffer TextureSizeBuffer : register(b0)
                    {
                        uint Width;
                        uint Height;
                    }
                    [numthreads(16, 16, 1)]
                    void main(uint3 dispatchThreadID : SV_DispatchThreadID)
                    {
                        uint2 texCoord = dispatchThreadID.xy;
                        if (texCoord.x < Width && texCoord.y < Height)
                        {
                            float4 color = InputTexture[texCoord];
                            float gray = (color.r + color.g + color.b) / 3.0;
                            OutputTexture[texCoord] = gray;
                        }
                    }";
                string compareShaderCode = @"
                    Texture2D<float> OldTexture : register(t0);
                    Texture2D<float> NewTexture : register(t1);
                    RWStructuredBuffer<uint> ResultBuffer : register(u0);
                    cbuffer TextureSizeBuffer : register(b0)
                    {
                        uint Width;
                        uint Height;
                    }
                    [numthreads(16, 16, 1)]
                    void main(uint3 dispatchThreadID : SV_DispatchThreadID)
                    {
                        uint2 texCoord = dispatchThreadID.xy;
                        if (texCoord.x < Width && texCoord.y < Height)
                        {
                            float oldPixel = OldTexture[texCoord];
                            float newPixel = NewTexture[texCoord];
                            float diff = abs(oldPixel - newPixel);
                            if (diff > 0.05)
                            {
                                InterlockedAdd(ResultBuffer[0], 1);
                            }
                        }
                    }";

                using (var grayscaleBytecode = SharpDX.D3DCompiler.ShaderBytecode.Compile(grayscaleShaderCode, "main", "cs_5_0"))
                using (var compareBytecode = SharpDX.D3DCompiler.ShaderBytecode.Compile(compareShaderCode, "main", "cs_5_0"))
                {
                    if (grayscaleBytecode.Bytecode != null && compareBytecode.Bytecode != null)
                    {
                        grayscaleComputeShader = new SharpDX.Direct3D11.ComputeShader(d3dDevice, grayscaleBytecode.Bytecode);
                        compareComputeShader = new SharpDX.Direct3D11.ComputeShader(d3dDevice, compareBytecode.Bytecode);
                        System.Console.WriteLine("Compute shaders compiled successfully.");
                    }
                    else
                    {
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
        /// </summary>
        public bool ProcessScreenshotOnGPU()
        {
            System.Console.WriteLine($"Screenshot taken at {System.DateTime.Now:HH:mm:ss.fff}");
            SharpDX.DXGI.Resource desktopResource = null;
            SharpDX.DXGI.OutputDuplicateFrameInformation frameInfo;
            bool frameAcquired = false;

            try
            {
                int retryCount = 0;
                const int maxRetries = 3;
                while (retryCount < maxRetries && !frameAcquired)
                {
                    try
                    {
                        System.Console.WriteLine("Attempting to acquire frame...");
                        outputDuplication.AcquireNextFrame(500, out frameInfo, out desktopResource);
                        frameAcquired = true;
                        System.Console.WriteLine("Frame acquired successfully.");
                    }
                    catch (SharpDX.SharpDXException ex) when (ex.HResult == unchecked((int)0x887A0001))
                    {
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

                using (var capturedTexture = desktopResource.QueryInterface<SharpDX.Direct3D11.Texture2D>())
                {
                    System.Console.WriteLine($"Captured texture dimensions: Width={capturedTexture.Description.Width}, Height={capturedTexture.Description.Height}, Format={capturedTexture.Description.Format}");

                    if (capturedTexture.Description.Width != currentGrayscaleTexture.Description.Width ||
                        capturedTexture.Description.Height != currentGrayscaleTexture.Description.Height)
                    {
                        System.Console.WriteLine($"Captured texture dimensions mismatch: Expected {currentGrayscaleTexture.Description.Width}x{currentGrayscaleTexture.Description.Height}, Got {capturedTexture.Description.Width}x{capturedTexture.Description.Height}");
                        return false;
                    }

                    System.Console.WriteLine("Converting to grayscale...");
                    ConvertToGrayscaleGPU(capturedTexture);

                    bool significantChange = false;
                    if (previousGrayscaleTexture != null)
                    {
                        System.Console.WriteLine("Comparing screenshots...");
                        int diffCount = CompareScreenshotsGPU();
                        int totalPixels = currentGrayscaleTexture.Description.Width * currentGrayscaleTexture.Description.Height;
                        double diffPercentage = (double)diffCount / totalPixels * 100.0;
                        System.Console.WriteLine($"CompareScreenshotsGPU ran at {System.DateTime.Now:HH:mm:ss.fff}, DiffCount: {diffCount}, Total Pixels: {totalPixels}, Percentage: {diffPercentage:F2}%");
                        significantChange = diffPercentage >= N_PIXEL_THRESH_PCT;
                    }
                    else
                    {
                        System.Console.WriteLine("Previous grayscale texture not available, skipping comparison.");
                    }

                    System.Console.WriteLine("Copying current grayscale texture to previous...");
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
                if (frameAcquired)
                {
                    try
                    {
                        System.Console.WriteLine("Releasing frame...");
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
        /// Converts a color texture to grayscale using a compute shader.
        /// </summary>
        private void ConvertToGrayscaleGPU(SharpDX.Direct3D11.Texture2D inputTexture)
        {
            if (inputTexture.Description.Format != SharpDX.DXGI.Format.B8G8R8A8_UNorm)
            {
                System.Console.WriteLine($"Invalid input texture format: Expected B8G8R8A8_UNorm, Got {inputTexture.Description.Format}");
                throw new System.InvalidOperationException("Input texture format must be B8G8R8A8_UNorm.");
            }

            var context = d3dDevice.ImmediateContext;

            // Unbind any previous resources
            context.ComputeShader.SetUnorderedAccessView(0, null);
            context.ComputeShader.SetShaderResource(0, null);
            context.ComputeShader.SetConstantBuffer(0, null);

            // Create shader resource view for input texture
            var inputSRV = new ShaderResourceView(d3dDevice, inputTexture, new ShaderResourceViewDescription
            {
                Format = Format.B8G8R8A8_UNorm,
                Dimension = ShaderResourceViewDimension.Texture2D,
                Texture2D = { MipLevels = 1, MostDetailedMip = 0 }
            });

            SharpDX.Direct3D11.Buffer textureSizeBuffer = null;
            try
            {
                // Create a properly sized constant buffer (must be multiple of 16 bytes)
                textureSizeBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, new BufferDescription
                {
                    SizeInBytes = 16, // 4 uint values = 16 bytes (minimum size for constant buffer)
                    Usage = ResourceUsage.Default,
                    BindFlags = BindFlags.ConstantBuffer,
                    CpuAccessFlags = CpuAccessFlags.None
                });

                // Prepare data: width, height, and padding to 16 bytes
                var textureSizeData = new[] {
            (uint)inputTexture.Description.Width,
            (uint)inputTexture.Description.Height,
            0u, 0u  // Padding to 16 bytes
        };

                // Update the constant buffer with our data
                context.UpdateSubresource(textureSizeData, textureSizeBuffer);
                System.Console.WriteLine("Constant buffer created and updated.");

                // Set the compute shader first
                context.ComputeShader.Set(grayscaleComputeShader);
                System.Console.WriteLine("Compute shader set.");

                // Then set resources
                context.ComputeShader.SetConstantBuffer(0, textureSizeBuffer);
                System.Console.WriteLine("Constant buffer bound.");

                context.ComputeShader.SetShaderResource(0, inputSRV);
                System.Console.WriteLine("Input SRV bound.");

                context.ComputeShader.SetUnorderedAccessView(0, grayscaleUAV);
                System.Console.WriteLine("Output UAV bound.");

                // Dispatch the compute shader
                int threadGroupX = (inputTexture.Description.Width + 15) / 16;
                int threadGroupY = (inputTexture.Description.Height + 15) / 16;
                System.Console.WriteLine($"Dispatching compute shader with thread groups: X={threadGroupX}, Y={threadGroupY}");
                context.Dispatch(threadGroupX, threadGroupY, 1);
                System.Console.WriteLine("Compute shader dispatched.");

                // Unbind resources
                context.ComputeShader.SetConstantBuffer(0, null);
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
                System.Console.WriteLine("Resources unbound after dispatch.");
            }
            catch (SharpDX.SharpDXException ex)
            {
                System.Console.WriteLine($"Compute shader setup failed: {ex.Message}");
                throw;
            }
            finally
            {
                inputSRV.Dispose();
                textureSizeBuffer?.Dispose();
            }
        }

        /// <summary>
        /// Compares the previous and current grayscale textures.
        /// </summary>
        private int CompareScreenshotsGPU()
        {
            var context = d3dDevice.ImmediateContext;

            // Reset the result buffer
            int zero = 0;
            context.UpdateSubresource(ref zero, resultBuffer);

            // Create a properly sized constant buffer
            using (var textureSizeBuffer = new SharpDX.Direct3D11.Buffer(d3dDevice, new BufferDescription
            {
                SizeInBytes = 16, // 4 uint values = 16 bytes
                Usage = ResourceUsage.Default,
                BindFlags = BindFlags.ConstantBuffer,
                CpuAccessFlags = CpuAccessFlags.None
            }))
            {
                // Prepare data with padding
                var textureSizeData = new[] {
            (uint)currentGrayscaleTexture.Description.Width,
            (uint)currentGrayscaleTexture.Description.Height,
            0u, 0u  // Padding to 16 bytes
        };

                // Update the constant buffer
                context.UpdateSubresource(textureSizeData, textureSizeBuffer);

                // Unbind previous resources
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetShaderResource(1, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
                context.ComputeShader.SetConstantBuffer(0, null);

                // Set shader and resources
                context.ComputeShader.Set(compareComputeShader);
                context.ComputeShader.SetConstantBuffer(0, textureSizeBuffer);
                context.ComputeShader.SetShaderResource(0, previousGrayscaleSRV);
                context.ComputeShader.SetShaderResource(1, currentGrayscaleSRV);
                context.ComputeShader.SetUnorderedAccessView(0, resultUAV);

                // Dispatch
                int threadGroupX = (currentGrayscaleTexture.Description.Width + 15) / 16;
                int threadGroupY = (currentGrayscaleTexture.Description.Height + 15) / 16;
                context.Dispatch(threadGroupX, threadGroupY, 1);

                // Unbind resources
                context.ComputeShader.SetConstantBuffer(0, null);
                context.ComputeShader.SetShaderResource(0, null);
                context.ComputeShader.SetShaderResource(1, null);
                context.ComputeShader.SetUnorderedAccessView(0, null);
            }

            // Get the results
            context.CopyResource(resultBuffer, resultStagingBuffer);
            var dataBox = context.MapSubresource(resultStagingBuffer, 0, MapMode.Read, SharpDX.Direct3D11.MapFlags.None);
            int diffCount = System.Runtime.InteropServices.Marshal.ReadInt32(dataBox.DataPointer);
            context.UnmapSubresource(resultStagingBuffer, 0);

            return diffCount;
        }

        /// <summary>
        /// Disposes of all DirectX resources.
        /// </summary>
        public void Dispose()
        {
            previousGrayscaleTexture?.Dispose();
            currentGrayscaleTexture?.Dispose();
            grayscaleUAV?.Dispose();
            previousGrayscaleSRV?.Dispose();
            currentGrayscaleSRV?.Dispose();
            resultUAV?.Dispose();
            resultBuffer?.Dispose();
            resultStagingBuffer?.Dispose();
            grayscaleComputeShader?.Dispose();
            compareComputeShader?.Dispose();
            outputDuplication?.Dispose();
            output?.Dispose();
            adapter?.Dispose();
            dxgiFactory?.Dispose();
            d3dDevice?.Dispose();
        }
    }
}