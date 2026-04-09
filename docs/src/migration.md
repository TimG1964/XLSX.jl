# Migration Guide

## Migrating from v0.10 to v0.11

There is only one breaking change in v0.11.0.

- `infer_eltypes` now defaults to `true` (e.g. in `gettable` and `readtable`). This is the more common use case but,
  if it is not *your* use case you will need explicitly to set `infer_eltypes = false` in the relevant functions.

All other changes either introduce new functionality (documented elsewhere) or relate to internals only.


## Migrating Legacy Code to v0.8+

!!! note

    The sections below were written as a guide to describe migrating from a pre v0.8 version 
    of XLSX.jl to v0.8. They are largely only historic now.

Version `v0.8` introduced a breaking change on methods [`XLSX.gettable`](@ref) and [`XLSX.readtable`](@ref).

These methods used to return a tuple `data, column_labels`.
On XLSX `v0.8` these methods return a `XLSX.DataTable` struct that implements [`Tables.jl`](https://github.com/JuliaData/Tables.jl) interface.

### Basic code replacement

Before

```julia
data, col_names = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
```

After

```julia
dtable = XLSX.readtable(joinpath(data_directory, "general.xlsx"), "table4")
data, col_names = dtable.data, dtable.column_labels
```

### Reading DataFrames

Since `XLSX.DataTable` implements `Tables.jl` interface,
the result of `XLSX.gettable` or `XLSX.readtable` can be
passed to a `DataFrame` constructor.

Before

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet")...)
```

After

```julia
df = DataFrame(XLSX.readtable("myfile.xlsx", "mysheet"))
```
