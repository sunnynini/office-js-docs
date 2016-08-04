# Custom XML Parts

_This branch is dedicated to describing the custom XML part design_

Microsoft Office documents allow users to embed custom XML data as part of the document. This data is named a custom XML part.

You can create and modify custom XML parts in a document by using the APIs described below. 

Usage: 

Most of the XML parts in the document are built-in parts that help to define the structure and the state of the document. However, documents can also contain custom XML parts, which you can use to store arbitrary XML data in the documents.

The XML file formats enable applications to extend the document by stroing add-in usage, state, or configuration data as part of the document. Since it travels with the document, it is very convenient to understand the state of the document. 

Current JavaScript API support: 

Currently, only Word supports custom XML part through the JavaScript 1.0 API set. It allows XML part and node level access with XML part level event support. 

Addition: 

We are adding custom XML part API to the modern 1.1+ API set for Word, Excel and PowerPoint. Excel and Word will get the new APIs first followed by PowerPoint. This open spec is about the new custom XML part API design.

## Workbook, Document Object 

Workbook and Document are the top level objects in Excel and Word respectively. A new custom XML parts collection will be made available at this top level. 

### Properties

Among other properties of this object, it will contain the custom XML part collection. 

| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|customXMLParts|CustomXMLPartCollection|Represents a collection of customXML parts associated with the workbook/document.|


## CustomXmlPartCollection Object 

A collection of custom XML parts.

## Properties

| Property	   | Type	|Description| 
|:---------------|:--------|:----------|
|items|CustomXmlPart|A collection of customXmlPart objects. Read-only.|


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[add(xml: string)](#addxml-string)|CustomXmlPart|Creates a new custom XML part and adds it to the collection.|
|[getByNamespace(namespaceUri: string)](#getbynamespacenamespaceuri-string)|CustomXmlPartCollection|Gets a new collection of custom XML parts whose namespaces match the given namespace.|
|[getItem(id: string)](#getitemid-string)|CustomXmlPart|Gets a custom XML part based on its ID.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### add(xml: string)
Creates a new custom XML part and adds it to the collection.

#### Syntax
```js
customXmlPartCollectionObject.add(xml);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xml|string|Custom XML content. Must be a valid XML fragment.|

#### Returns
CustomXmlPart

### getByNamespace(namespaceUri: string)
Gets a new collection of custom XML parts whose namespaces match the given namespace.

#### Syntax
```js
customXmlPartCollectionObject.getByNamespace(namespaceUri);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|namespaceUri|string||

#### Returns
CustomXmlPartCollection

### getItem(id: string)
Gets a custom XML part based on its ID.

#### Syntax
```js
customXmlPartCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|string|ID of the object to be retrieved.|

#### Returns
CustomXmlPart


## CustomXML Part object

Represents a custom XML part object in a workbook/document.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|builtIn|bool|Indicates whether the custom XML part is built-in. Read-only.|
|id|string|The custom XML part's ID. Read-only.|
|namespaceUri|string|The custom XML part's namespace URI. Read-only.|


## Relationships
| Relationship | Type	|Description| 
|:---------------|:--------|:----------|
|xml|String|The custom XML part's XML content.|


## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|delete|void|Deletes the custom XML part.|

### Method Details

### delete()
Deletes the custom XML part.

#### Syntax
```js
customXmlPartObject.delete();
```

#### Parameters
None

#### Returns
void

## Node level access

There are two options for providing the node level access.  

One is to define node level strongly typed objects as part of the custom XML object.  

* `customXmlnodes` collection available as part of the custom XML part. 
* Basic node details (`baseName`, `nameSpaceUri`, etc.) and methods: get node, set node, remove node, etc. made available on the custom XML node. This matches closely with the current API support in Word.

The second option is to make the XML string available for usage with open source libraries such as `xml2js`. Such libraries provide node level access (read/write) and many other advanced XML reading, writing options that add-ins may need anyway. 


## Event support 

We'll detail the event support for custom XML in the eventing design spec. 

| Event type| Description|
|:------|:------|
|Data node deleted|	Occurs when a node is deleted.|
|Data node inserted|	Occurs when a node is inserted.|
|Data node replaced|	Occurs when a node is replaced.|

**[Tell us what you think](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-CXP)**






