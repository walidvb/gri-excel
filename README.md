# GRI Excel Creator

This micro-service only creates and returns an excel file built based on the template.

## Request expected is 

POST /

```
"data": {
    "id": 102,
    "title": "the one",
    "slug": "the-one",
    "provider_id": 1,
    "note": null,
    "client": "{}",
    "provider": {
      "id": 1,
      "name": "FelipeCorp",
      "created_at": null,
      "updated_at": null,
      "deleted_at": null
    },
    "created_at": "2019-07-31T19:50:17.000000Z",
    "updated_at": "2019-07-31T19:59:16.000000Z",
    "version": {
      "project_id": 102,
      "version_reference": 1,
      "rooms": [{
        "name": "Travaux G\u00e9n\u00e9raux",
        "slug": "Travaux-Generaux",
        "steps": [{
          "id": 1,
          "tags": [],
          "unit": "m3",
          "price": "3.00",
          "category": "wood",
          "quantity": "2",
          "room_type": "general",
          "unit_price": "1.90",
          "description": "Soluta a ipsa qui quis."
        },
          {
            "id": 1,
            "tags": [],
            "unit": "m3",
            "price": "3.00",
            "category": "wood",
            "quantity": "2",
            "room_type": "general",
            "unit_price": "1.90",
            "description": "Soluta a ipsa qui quis."
          },
          {
            "id": 1,
            "tags": [],
            "unit": "m3",
            "price": "3.00",
            "category": "wood",
            "quantity": "2",
            "room_type": "general",
            "unit_price": "1.90",
            "description": "Soluta a ipsa qui quis."
          }],
        "roomType": "general",
        "stepCompleted": [true, true, null]
      }, {
        "name": "test",
        "note": "13",
        "slug": "test",
        "steps": [{
          "id": 2,
          "tags": [],
          "unit": "m2",
          "price": "2.70",
          "category": "floor",
          "quantity": "3",
          "room_type": "dry",
          "unit_price": "1.20",
          "description": "Ducimus aut veniam numquam in architecto neque."
        }],
        "roomType": "dry",
        "stepCompleted": [true, true, 2]
      }],
      "locked": 0,
      "price": "14.10",
      "note": null,
      "rooms_count": 2,
      "xls_path": null,
      "pdf_path": null,
      "created_at": "2019-07-31T19:59:16.000000Z",
      "updated_at": "2019-07-31T23:45:11.000000Z"
    },
```