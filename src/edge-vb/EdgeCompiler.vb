Option Strict On

Imports Microsoft.VisualBasic
Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks

Public Class EdgeCompiler
	Shared ReadOnly referenceRegex As New Regex("^[\ \t]*(?:\/{2})?\#r[\ \t]+""([^""]+)""", RegexOptions.Multiline)
	Shared ReadOnly usingRegex As New Regex("^[\ \t]*(using[\ \t]+[^\ \t]+[\ \t]*\;)", RegexOptions.Multiline)
    Shared ReadOnly debuggingEnabled As Boolean = Not String.IsNullOrEmpty(Environment.GetEnvironmentVariable("EDGE_VB_DEBUG"))
    Shared ReadOnly debuggingSelfEnabled As Boolean = Not String.IsNullOrEmpty(Environment.GetEnvironmentVariable("EDGE_VB_DEBUG_SELF"))
    Shared ReadOnly cacheEnabled As Boolean = Not String.IsNullOrEmpty(Environment.GetEnvironmentVariable("EDGE_VB_CACHE"))
    Shared referencedAssemblies As New Dictionary(Of String, Dictionary(Of String, Assembly))()
	Shared funcCache As New Dictionary(Of String, Func(Of Object, Task(Of Object)))()

	Shared Sub New()
		AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve
	End Sub

	Private Shared Function CurrentDomain_AssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
		Dim result As Assembly = Nothing
		Dim requesting As Dictionary(Of String, Assembly)
		If referencedAssemblies.TryGetValue(args.RequestingAssembly.FullName, requesting) Then
			requesting.TryGetValue(args.Name, result)
		End If

		Return result
	End Function

	Public Function CompileFunc(parameters As IDictionary(Of String, Object)) As Func(Of Object, Task(Of Object))
		Dim source As String = DirectCast(parameters("source"), String)
		Dim lineDirective As String = String.Empty
		Dim fileName As String = Nothing
		Dim lineNumber As Integer = 1

        ' read source from file
        If source.EndsWith(".vb", StringComparison.InvariantCultureIgnoreCase) OrElse source.EndsWith(".vbx", StringComparison.InvariantCultureIgnoreCase) Then
            ' retain fileName for debugging purposes
            If debuggingEnabled Then
                fileName = source
            End If

            source = File.ReadAllText(source)
        End If

        If debuggingSelfEnabled Then
			Console.WriteLine("Func cache size: " & funcCache.Count)
		End If

        Dim originalSource As String = source
        If funcCache.ContainsKey(originalSource.ToString()) Then
            If debuggingSelfEnabled Then
                Console.WriteLine("Serving func from cache.")
            End If

            Return funcCache(originalSource.ToString())
        ElseIf debuggingSelfEnabled Then
            Console.WriteLine("Func not found in cache. Compiling.")
		End If

		' add assembly references provided explicitly through parameters
		Dim references As New List(Of String)()
		Dim v As Object
		If parameters.TryGetValue("references", v) Then
			For Each reference As Object In DirectCast(v, Object())
				references.Add(DirectCast(reference, String))
			Next
		End If

		' add assembly references provided in code as [//]#r "assemblyname" lines
		Dim match As Match = referenceRegex.Match(source)
		While match.Success
			references.Add(match.Groups(1).Value)
			source = source.Substring(0, match.Index) & source.Substring(match.Index + match.Length)
			match = referenceRegex.Match(source)
		End While

		If debuggingEnabled Then
			Dim jsFileName As Object
			If parameters.TryGetValue("jsFileName", jsFileName) Then
				fileName = DirectCast(jsFileName, String)
				lineNumber = CInt(parameters("jsLineNumber"))
			End If

			If Not String.IsNullOrEmpty(fileName) Then
				lineDirective = String.Format("#line {0} ""{1}""" & vbLf, lineNumber, fileName)
			End If
		End If

		' try to compile source code as a class library
		Dim assembly__1 As Assembly
		Dim errorsClass As String
		If Not Me.TryCompile(lineDirective & source, references, errorsClass, assembly__1) Then
			' try to compile source code as an async lambda expression

			' extract using statements first
			Dim usings As String = ""
			match = usingRegex.Match(source)
			While match.Success
				usings += match.Groups(1).Value
				source = source.Substring(0, match.Index) & source.Substring(match.Index + match.Length)
				match = usingRegex.Match(source)
			End While

			Dim errorsLambda As String
			source = usings & "using System;" & vbLf & "using System.Threading.Tasks;" & vbLf & "public class Startup {" & vbLf & "    public async Task<object> Invoke(object ___input) {" & vbLf & lineDirective & "        Func<object, Task<object>> func = " & source & ";" & vbLf & "#line hidden" & vbLf & "        return await func(___input);" & vbLf & "    }" & vbLf & "}"

			If debuggingSelfEnabled Then
                Console.WriteLine("Edge-vb trying to compile async lambda expression:")
                Console.WriteLine(source)
			End If

			If Not TryCompile(source, references, errorsLambda, assembly__1) Then
                Throw New InvalidOperationException("Unable to compile VB code." & vbLf & "----> Errors when compiling as a CLR library:" & vbLf & errorsClass & vbLf & "----> Errors when compiling as a CLR async lambda expression:" & vbLf & errorsLambda)
            End If
		End If

		' store referenced assemblies to help resolve them at runtime from AppDomain.AssemblyResolve
		referencedAssemblies(assembly__1.FullName) = New Dictionary(Of String, Assembly)()
        For Each reference As String In references
            Try
                Dim referencedAssembly As Assembly = Assembly.UnsafeLoadFrom(reference)
                referencedAssemblies(assembly__1.FullName)(referencedAssembly.FullName) = referencedAssembly
                ' empty - best effort
            Catch
            End Try
        Next

        ' extract the entry point to a class method
        Dim startupType As Type = assembly__1.[GetType](DirectCast(parameters("typeName"), String), True, True)
		Dim instance As Object = Activator.CreateInstance(startupType, False)
		Dim invokeMethod As MethodInfo = startupType.GetMethod(DirectCast(parameters("methodName"), String), BindingFlags.Instance Or BindingFlags.[Public])
		If invokeMethod Is Nothing Then
			Throw New InvalidOperationException("Unable to access CLR method to wrap through reflection. Make sure it is a public instance method.")
		End If

		' create a Func<object,Task<object>> delegate around the method invocation using reflection
		Dim result As Func(Of Object, Task(Of Object)) = Function(input) 
		Return DirectCast(invokeMethod.Invoke(instance, New Object() {input}), Task(Of Object))

End Function

		If cacheEnabled Then
			funcCache(originalSource) = result
		End If

		Return result
	End Function

	Private Function TryCompile(source As String, references As List(Of String), ByRef errors As String, ByRef assembly As Assembly) As Boolean
		Dim result As Boolean = False
		assembly = Nothing
		errors = Nothing

		Dim options As New Dictionary(Of String, String)() From { _
			{"CompilerVersion", "v4.0"} _
		}
        Dim vbc As New VBCodeProvider(options)
        Dim parameters As New CompilerParameters()
        parameters.GenerateInMemory = True
        parameters.IncludeDebugInformation = debuggingEnabled
        parameters.ReferencedAssemblies.AddRange(references.ToArray())
        parameters.ReferencedAssemblies.Add("System.dll")
        parameters.ReferencedAssemblies.Add("System.Core.dll")
        parameters.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")
        If Not String.IsNullOrEmpty(Environment.GetEnvironmentVariable("EDGE_VB_TEMP_DIR")) Then
            parameters.TempFiles = New TempFileCollection(Environment.GetEnvironmentVariable("EDGE_VB_TEMP_DIR"))
        End If
        Dim results As CompilerResults = vbc.CompileAssemblyFromSource(parameters, source)
        If results.Errors.HasErrors Then
			For Each [error] As CompilerError In results.Errors
				If errors Is Nothing Then
					errors = [error].ToString()
				Else
					errors += vbLf & [error].ToString()
				End If
			Next
		Else
			assembly = results.CompiledAssembly
			result = True
		End If

		Return result
	End Function
End Class
