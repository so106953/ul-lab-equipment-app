# Prototype Instructions

Run the local server yourself and open the preview in the browser available to this environment. Do not give the user server-start instructions when you can run it.

Before making substantial visual changes, use the Product Design plugin's `get-context` skill when the visual source is unclear or no longer matches the current goal. When the user gives durable prototype-specific design feedback, preferences, or decisions, record them in `AGENTS.md`.

When implementing from a selected generated mock, treat that image as the source of truth for layout, component anatomy, density, spacing, color, typography, visible content, and hierarchy.

## Confirmed design direction

- Red UL Solutions visual language: supplied UL Solutions logo asset, red app shell, white operational surfaces.
- Equipment-centric lending/return workspace with a compact sidebar, equipment hero, numbered form sections, and fixed transaction summary.
- Submission confirmation must use the same lightweight fireworks effect on desktop and mobile; it must never obscure the workspace with a blank screen or delay the form after a cloud write.
- Record lists must show a searchable equipment number and keep issue and return history in separate navigation views. An active loan blocks a second issue of the same equipment number, while still allowing its return registration.
- Equipment availability and normal status are released only by a later return record with the same equipment name, group, and equipment number; returning a different unit must not release an active loan.
- A return requires the active borrower's exact staff ID for that same equipment number. Mismatched staff IDs must be blocked with a clear correction message.
