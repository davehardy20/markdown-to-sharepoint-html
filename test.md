# Test Document for Markdown to SharePoint HTML Conversion

This document tests various markdown elements and code blocks.

## Headings

### Level 3 Heading
#### Level 4 Heading
##### Level 5 Heading
###### Level 6 Heading

## Text Formatting

**Bold text** and *italic text* and ***bold italic***

~~Strikethrough~~ text is also supported.

## Lists

Unordered list:
- Item 1
- Item 2
  - Nested item 2a
  - Nested item 2b
- Item 3

Ordered list:
1. First item
2. Second item
3. Third item

## Links and Images

Visit [GitHub](https://github.com) for more information.

![Sample Image](https://via.placeholder.com/150)

## Blockquotes

> This is a blockquote.
> It can span multiple lines.

## Code Blocks

### JavaScript Example

```javascript
function greet(name) {
  console.log(`Hello, ${name}!`);
}

greet('World');
```

### Python Example

```python
def calculate_sum(a, b):
    return a + b

result = calculate_sum(5, 3)
print(f"Result: {result}")
```

### Bash Example

```bash
#!/bin/bash

echo "Starting deployment..."
cd /app
npm install
npm start
```

### PowerShell Example

```powershell
# Install a module
Install-Module -Name Az -Force

# Connect to Azure
Connect-AzAccount

# List resources
Get-AzResourceGroup
```

### JSON Example

```json
{
  "name": "test",
  "version": "1.0.0",
  "dependencies": {
    "markdown-it": "^14.0.0",
    "highlight.js": "^11.0.0"
  }
}
```

### SQL Example

```sql
SELECT
    id,
    name,
    email,
    created_at
FROM users
WHERE status = 'active'
ORDER BY created_at DESC
LIMIT 10;
```

### HTML Example

```html
<!DOCTYPE html>
<html>
<head>
  <title>Sample Page</title>
</head>
<body>
  <h1>Welcome</h1>
  <p>This is a sample HTML page.</p>
</body>
</html>
```

## Inline Code

Use `npm install` to install dependencies.

## Tables

| Name | Type | Status |
|------|------|--------|
| Task 1 | High | Done |
| Task 2 | Medium | In Progress |
| Task 3 | Low | Pending |

## Horizontal Rule

---

## Mixed Content Example

This paragraph has **bold**, *italic*, `inline code`, and a [link](https://example.com).

### Code Block with Multiple Languages

```yaml
version: '3.8'
services:
  web:
    build: .
    ports:
      - "3000:3000"
```

> Note: YAML files are commonly used for configuration.

---

## Complex Code Example

```typescript
interface User {
  id: number;
  name: string;
  email: string;
}

class UserService {
  private users: User[] = [];

  addUser(user: User): void {
    this.users.push(user);
  }

  getUserById(id: number): User | undefined {
    return this.users.find(u => u.id === id);
  }
}

const service = new UserService();
service.addUser({ id: 1, name: 'John', email: 'john@example.com' });
```

## Test Summary

This document tests:
- All heading levels (H1-H6)
- Bold, italic, strikethrough
- Ordered and unordered lists
- Links and images
- Blockquotes
- Code blocks with syntax highlighting
- Inline code
- Tables
- Horizontal rules
- Mixed formatting
