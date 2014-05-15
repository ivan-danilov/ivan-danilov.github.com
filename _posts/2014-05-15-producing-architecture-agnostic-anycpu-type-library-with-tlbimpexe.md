---
layout: post
title: "Producing architecture agnostic (AnyCPU) type library with tlbimp.exe"
description: ""
category: 
tags: [com, interop]
---
{% include JB/setup %}


While working on [UIA customization in managed code](https://github.com/ivan-danilov/uia-custom-pattern-managed), I've encountered issue with desired architecture: I know that Windows provide UIA for both x86 and x64, and I can make them both work... but only as a two separate architecture-specific assemblies. I wanted to have AnyCPU assembly with pure MSIL inside, that can be loaded into either x86 or x64 process without issues.

Let me describe briefly how COM components could be consumed from the .NET world.

Native COM interface description language is IDL - Microsoft flavor of Interface Definition Language. It resembles C header file. For example, here is excerpt from UIAutomationCore.idl (you can find it in the Windows SDK):

```
//
//  IRawElementProviderSimple
//
[object, uuid(d6dd68d1-86fd-4332-8666-9abedea2d24c), pointer_default(unique)]
interface IRawElementProviderSimple : IUnknown
{
    [propget] HRESULT ProviderOptions (
        [out, retval] enum ProviderOptions * pRetVal );

    HRESULT GetPatternProvider (
        [in] PATTERNID patternId,
        [out, retval] IUnknown ** pRetVal );

    HRESULT GetPropertyValue (
        [in] PROPERTYID propertyId,
        [out, retval] VARIANT * pRetVal );

    [propget] HRESULT HostRawElementProvider (
        [out, retval] IRawElementProviderSimple ** pRetVal );
}
```

Now, we can take that interface and write our own wrapper interface, decorate it with `ComVisible`, `InterfaceType` and `GUID` attributes, so that .NET runtime/marshaler is able to consume it and create CCW/RCW correctly. But it is tedious and error-prone (hey, there's not a single interface!). If you're interested, [here](http://blog.kutulu.org/2012/01/com-interop-part-10-recapping-and.html) is the marvelous introduction how to write managed wrappers for COM. It is very brief - just 10 posts ;)  

Fortunately, there's a shorter way (not necessary better, though). Microsoft provides utility called tlbimp.exe that is able to produce managed wrappers automatically, if given binary Type Library file (*.tlb). Where could you get this tlb? Well, *.tlb is the same to *.idl what *.exe is to *.cs - it is compiled form. And there's a compiler called Microsoft IDL compiler or midl.exe which takes *.idl and produces *.tlb.

In principle, it is simple: we take UIAutomationCore.idl, pass it through midl.exe, get UIAutomationCore.tlb, feed it into tlbimp.exe and get managed UIAutomationCore.dll with wrappers.

But, there's some problems. It does not always make translation correctly, because IDL doesn't provide all required information. For example, function may take pointer and integer, and tlbimp.exe has no way to know that is it C-style array where pointer points to first element of an array, and integer denotes array size. For such cases there's separate [page](http://msdn.microsoft.com/en-us/library/ek1fb3c6%28v=vs.110%29.aspx) exists.

So, in order to get better results, we have to change result of tlbimp.exe a bit. How? First, ildasm.exe, then change as needed and then ilasm.exe results back. That's precisely the steps made by [these CustomBuild actions](https://github.com/ivan-danilov/uia-custom-pattern-managed/blob/master/UiaCoreInterop/UIACoreInterop.vcxproj#L39):
1. IDL -> midl -> TLB
2. TLB -> tlbimp -> managed DLL
3. managed DLL -> ildasm -> IL code
4. IL code -> our console project that just plainly replaces certain parts of IL code -> modified IL
5. modified IL -> ilasm -> final managed DLL

So, where is architecture comes in? COM is able to service even situation where calls come from x64 process and client is written in x86 (or vice-versa) - with help of so-called DLL Surrogates, what's the problem could we have here? It turns out, it is another place, where tlbimp.exe behaves not correctly. When you're importing tlb from Visual Studio IDE, you can't import it as AnyCPU and have to specify concrete architecture (VS just calls tlbimp.exe under the cover). But tlbimp.exe has `/machine:` switch, which can be `x64`, `x86`, or `Agnostic`. Yes, you've guessed correctly, if you specify `/machine:Agnostic` - resulting dll would have Platform=AnyCPU.

But in case of UIAutomationCore, it won't work. x64 and x86 work beautifully and Agnostic is not. To be precise, it works under x86, but refuses to work under x64 process (it loads fine, without `BadImageFormatException`, but still doesn't work).

I've tried to analyze the difference between dlls produced with different `/machine` switch values... and found that the only difference is the `.pack 4` in x86 and Agnostic cases, and `.pack 8` in x64 case. Obviously, it is structure alignment in memory. Hey, but why it produces `.pack 4` for AnyCPU?! Good question. I'd like to know as well. Option `.pack 0`, which means "use the most appropriate alignment for the platform you're executing on" seems much more logical from tlbimp's point of view... It'd result in `.pack 4` on x86 and `.pack 8` on x64, which is what's expected if IDL doesn't have `pragma` directives to explicitly state alignment.

Finally, after I added some more [customizations](https://github.com/ivan-danilov/uia-custom-pattern-managed/blob/master/CustomizeUiaInterop/Program.cs#L88) into our console project (see step 4 above) I was able to make all my UIA customization assemblies AnyCPU, and it not only loads successfully into address space of any process, but also works :) 