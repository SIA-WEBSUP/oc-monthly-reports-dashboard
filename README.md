# oc-monthly-reports

Monorepo for O&C monthly report tooling.

## Structure

- apps/master-dashboard: Google Sheet-bound master script
  - Script ID: 1PELivo8WXRYdPXuNnK_BnTPIEszz4SkWf6MH5XptIdVzvLwAUZinKGg7
- apps/oc-doc-bound-script: Google Doc-bound script (menu + library calls)
  - Script ID: 1JfO4fZIJftid3MmVNL5TR6AMndcz-HVQnD3tI-4gEJwFVTUXmJ6rkhjF
- apps/oc-library: Shared Apps Script library
  - Script ID: 1JE4hStZsGvGD7xYv63_ohZKem7zPKYS6Gvf8HLcbBk1PfQLQmXJbrPeK

## Bootstrap

Use clasp in each app folder:

- cd apps/master-dashboard && clasp pull
- cd apps/oc-doc-bound-script && clasp clone 1JfO4fZIJftid3MmVNL5TR6AMndcz-HVQnD3tI-4gEJwFVTUXmJ6rkhjF
- cd apps/oc-library && clasp clone 1JE4hStZsGvGD7xYv63_ohZKem7zPKYS6Gvf8HLcbBk1PfQLQmXJbrPeK
