# Cache Utilties

The cache worksheet functions provide direct ways to interact with the XlBlocks cache. The cache is a key-value store that allows you to store and retrieve data efficiently. Cache handles are hashed using the address of the cell they are located in. Thus, no matter the type of the object that is being cached, the handle will remain the same so long as it is located in the same cell on the same workbook.

In most cases, the cache functions documented below will not be used directly. Instead, functions that return XlBlocks objects will automatically cache the results, and functions that have XlBlocks objects in their parameters will expect cache handles. However, the `XBCache_Handle` function is an exception to this rul. As it returns the cache handle for the cell it is located in, it can be used in a pattern such as `=IF(toggle, SOME_FUNCTION(), XBCache_Handle())`, where `SOME_FUNCTION()` is a long-running calculation (such as loading a large csv into a table). This allows one to access the value associated with the results of that function on repeated calculations without having to actually rerun the long-running calculation itself. 

---

## Cache Functions
