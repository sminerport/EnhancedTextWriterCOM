# EnhancedTextWriterCOM

## Overview

EnhancedTextWriterCOM is a comprehensive example project that demonstrates how to create a COM object in C#, register it for use in VBA (Visual Basic for Applications), and showcase its functionality through early and late binding techniques. This project is perfect for developers looking to extend the capabilities of Microsoft Office applications, such as Word or Excel, by integrating .NET functionalities.

For a detailed step-by-step tutorial on creating COM objects in .NET and using them in VBA, please visit the following links:

- [Creating COM Objects in .NET](https://scottminer.netlify.app/post/creating-com-objects-in-dotnet/)

## Project Structure

The repository contains the following key components:

- **C# Project (`EnhancedTextWriterLib`)**: The core of the project where the COM object (`EnhancedTextWriter`) is created. This project includes the implementation of the `SaveToFile` method, which allows text from a VBA form to be saved to a file.
  
- **DOCM File (`EnhancedTextWriter.docm`)**: A Microsoft Word macro-enabled document that contains the VBA code necessary to interact with the COM object. The DOCM file includes examples of both early and late binding in VBA.

- **VBA Code (`VBA/EnhancedTextWriterVBA.bas`)**: The VBA module that demonstrates how to use the `EnhancedTextWriter` COM object within a Word document. This code shows the setup for both early and late binding scenarios.

## Features

- **Early Binding**: Demonstrates how to set up early binding in VBA by referencing the COM object directly, providing IntelliSense support and compile-time type checking.

- **Late Binding**: Illustrates how to use late binding in VBA, creating the COM object at runtime with the `CreateObject` function for more flexibility.

- **COM Interoperability**: Shows how to expose .NET functionality to VBA by registering a .NET assembly as a COM object.

## Getting Started

### Prerequisites

- Visual Studio 2019 or later with .NET Framework 4.8 or later installed.
- Microsoft Word 2016 or later with support for macro-enabled documents (DOCM).
- Basic understanding of C# and VBA.

### Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/sminerport/EnhancedTextWriterCOM.git
   ```
2. **Build the C# Project**:
   - Open the `EnhancedTextWriterLib` project in Visual Studio.
   - Build the solution to compile the DLL.

3. **Register the COM Object**:
   - Use the `regasm` tool to register the compiled DLL for COM interop:
     ```bash
     regasm /codebase /tlb EnhancedTextWriterLib.dll
     ```

4. **Open the Word Document**:
   - Open the `EnhancedTextWriter.docm` file in Microsoft Word.
   - Enable macros if prompted.

5. **Run the VBA Examples**:
   - Access the VBA editor in Word (`Alt + F11`).
   - Explore the `EnhancedTextWriterVBA` module to see examples of early and late binding.
   - Run the VBA code to interact with the COM object.

## Usage

This project is intended as a learning resource for developers who want to:

- Understand how to create and expose COM objects in .NET.
- Learn the differences between early and late binding in VBA.
- Extend the functionality of Office applications using .NET and VBA.

## Contributing

Contributions are welcome! If you have suggestions, improvements, or bug fixes, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or further information, feel free to reach out via GitHub issues or contact me directly at [scott.miner.data.scientist@gmail.com](mailto:scott.miner.data.scientist@gmail.com).

---

Happy coding!
