# SportsClubMembershipSystem
A demo on how to create and maintain membership details for a sports club. This software component validates input.
========
SCENARIO
========
A: provides an outline design specification for a software component to validate
input.

B. Criteria for design
Field and validation requirements include the following:
1. Membership Number
- Not empty
- Modulus 11
- Must be 6 digits
- Numeric
2. First Name
- None
3. Last Name 
- None
4. Address 
- None
5. Postcode 
- None
6. Sex
- M or F (upper or lower allowed)
7. Date of Birth 
- dd/mm/yyyy Full date check
8. Join Date 
- dd/mm/yyyy Full date check
9. Type of Membership
- F, S, T or B (upper or lower allowed)
10. Subscription Due Month 
- MMM e.g. Jan

=============
REQUIREMENTS
=============
1. In order to build THIS project, you will still need Visual Basic 6 installed.
or
2. Download: https://visualstudiogallery.msdn.microsoft.com/0abaccb5-76a1-4022-9e0e-f6832c621162/file/122982/1/VisualBasic6X.vsix
This extension allows you to load and build VB6 projects from within Visual Studio.

How To:
-You will need to use the project converter to convert the project from VBP to VBPX.
-The project converter is available from the project homepage:
 +Use the converter to convert your VBP to VBPX
 +Add your VBPX project to your Visual Studio solution.

Details:
The extension simply defines a new project type, and will turn your VB6 project into an MSBuild project file that can be fully integrated into the Visual Studio / MSBuild build process like C# and VB.NET projects.

Limitations:
It's recommended you still use VB6 to open, edit and debug your projects, as the language is not supported from within the IDE. This extension is purely to bring your VB6 projects into the Visual Studio buid process.

==============
PROJECT FILES
==============
1. MembershipSystem.vbp
2. MembershipSystem.vbw
3. MSSCCPRJ.SCC
4. SportsClubMembership.frm
5. SportsClubMembership.frx

