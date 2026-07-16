"""Seed a starter menu template so the admin doesn't start from a blank page.

Per spec section 6.2/9.2: prices here are conservative *placeholders*, not
verified store data. The admin must confirm every item/topping price against
the actual foodpanda page for the current week before publishing.
"""

from models import MenuTemplate, TemplateItem, TemplateTopping, db

VERIFY_NOTE = (
    "⚠️ 此範本價格為系統預設參考值，非即時店家資料。"
    "發布本週菜單前，請務必對照 foodpanda／店家頁面核對外送價與加料是否收費，"
    "再手動調整下方品項與加料金額。"
)

DEFAULT_TEMPLATES = [
    {
        "name": "50嵐 (同安店)",
        "foodpanda_link": "",
        "notes": VERIFY_NOTE,
        "items": [
            ("四季春", 30),
            ("烏龍綠", 30),
            ("阿薩姆紅茶", 30),
            ("蜂蜜緑茶", 35),
            ("冬瓜檸檬", 35),
            ("珍珠奶茶", 45),
            ("波霸奶茶", 45),
            ("紅茶拿鐵", 50),
            ("烏龍拿鐵", 50),
        ],
        "toppings": [
            ("珍珠", 5),
            ("波霸", 5),
            ("椰果", 5),
            ("仙草", 5),
            ("布丁", 5),
            ("燕麥", 5),
            ("寒天", 5),
        ],
    },
]


def seed_default_templates():
    for template_def in DEFAULT_TEMPLATES:
        exists = MenuTemplate.query.filter_by(name=template_def["name"]).first()
        if exists:
            continue
        template = MenuTemplate(
            name=template_def["name"],
            foodpanda_link=template_def["foodpanda_link"],
            notes=template_def["notes"],
        )
        db.session.add(template)
        db.session.flush()
        for order_idx, (name, price) in enumerate(template_def["items"]):
            db.session.add(
                TemplateItem(
                    template_id=template.id, name=name, price=price, sort_order=order_idx
                )
            )
        for order_idx, (name, price) in enumerate(template_def["toppings"]):
            db.session.add(
                TemplateTopping(
                    template_id=template.id, name=name, price=price, sort_order=order_idx
                )
            )
    db.session.commit()
