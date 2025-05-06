# IIS_Manager

A .NET library for automating and managing IIS websites, application pools, FTP directories, and virtual applications.  
Designed for use in deployment tools, admin utilities, or internal devops environments.

> ✅ Supports both legacy IIS DirectoryEntry (`IIS://`) and modern Microsoft.Web.Administration APIs.

---

## 🔧 Features

- Create, start, stop, or remove IIS websites and app pools
- Add host headers and set bindings
- Create virtual directories and virtual applications
- Manage FTP directories via IIS metabase
- Retrieve IIS site information as XML or DataSet
- Assign app pools to sites and virtual directories
- Modify NTFS folder permissions for web users
- Create local Windows users (optional use)

---

## 🚀 Getting Started

### Requirements

- .NET Framework 4.7.2+ or compatible
- Administrative privileges (required to interact with IIS)
- IIS must be installed on the target system

### Installation

Install the DLL into your project:

```bash
dotnet add reference path/to/IIS_Manager.dll
