---
name: clarity-first_coding-style
description: Enforces requirement clarity, structured modular coding, clean architecture, and consistent naming conventions.
---

# Clarity First Coding Style

## 🔎 Requirement Understanding Rule (CRITICAL)

- NEVER start coding immediately.
- FIRST fully understand the user's prompt and requirements.
- If anything is unclear, ASK clarifying questions before proceeding.
- Confirm assumptions when necessary.
- Do not partially guess requirements.

Coding should only begin after requirements are completely clear.

---

## 📁 Naming Conventions

### Folders
- Use **PascalCase**
- Example:
  - `UserManagement`
  - `PaymentGateway`

### Files
- Use **snake_case**
- Lowercase only
- Example:
  - `user_service.js`
  - `payment_handler.py`

---

## 🧠 Function Design Principles

- Always include a **one-line description** above every function.
- Large functions must be split into **small reusable functions**.
- Functions should have a **single responsibility**.
- Avoid nested statements whenever possible.
- Keep logic flat and readable.

Example:

```js
// Calculates total price including tax
function calculate_total_price(items) {}
