from datetime import datetime

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

SUGAR_OPTIONS = ["正常糖", "少糖", "半糖", "微糖", "無糖"]
ICE_OPTIONS = ["正常冰", "少冰", "微冰", "去冰", "溫", "熱"]


class Week(db.Model):
    __tablename__ = "weeks"

    id = db.Column(db.Integer, primary_key=True)
    week_key = db.Column(db.String(32), unique=True, nullable=False)
    status = db.Column(db.String(16), nullable=False, default="active")  # active | archived
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    menu = db.relationship(
        "Menu", backref="week", uselist=False, cascade="all, delete-orphan"
    )
    orders = db.relationship(
        "Order",
        backref="week",
        cascade="all, delete-orphan",
        order_by="Order.created_at",
    )
    settlement = db.relationship(
        "Settlement", backref="week", uselist=False, cascade="all, delete-orphan"
    )

    @property
    def grand_total_estimate(self):
        return sum(o.unit_price * o.qty for o in self.orders)


class Menu(db.Model):
    __tablename__ = "menus"

    id = db.Column(db.Integer, primary_key=True)
    week_id = db.Column(db.Integer, db.ForeignKey("weeks.id"), nullable=False, unique=True)
    store = db.Column(db.String(120), nullable=False, default="")
    foodpanda_link = db.Column(db.String(500), nullable=False, default="")
    deadline = db.Column(db.DateTime, nullable=True)
    notes = db.Column(db.Text, nullable=False, default="")
    published = db.Column(db.Boolean, nullable=False, default=False)

    items = db.relationship(
        "MenuItem",
        backref="menu",
        cascade="all, delete-orphan",
        order_by="MenuItem.sort_order",
    )
    toppings = db.relationship(
        "MenuTopping",
        backref="menu",
        cascade="all, delete-orphan",
        order_by="MenuTopping.sort_order",
    )

    @property
    def is_foodpanda(self):
        return "foodpanda" in (self.foodpanda_link or "").lower()


class MenuItem(db.Model):
    __tablename__ = "menu_items"

    id = db.Column(db.Integer, primary_key=True)
    menu_id = db.Column(db.Integer, db.ForeignKey("menus.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    price = db.Column(db.Integer, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)


class MenuTopping(db.Model):
    __tablename__ = "menu_toppings"

    id = db.Column(db.Integer, primary_key=True)
    menu_id = db.Column(db.Integer, db.ForeignKey("menus.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    price = db.Column(db.Integer, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)


class Order(db.Model):
    __tablename__ = "orders"

    id = db.Column(db.Integer, primary_key=True)
    week_id = db.Column(db.Integer, db.ForeignKey("weeks.id"), nullable=False)
    person_name = db.Column(db.String(80), nullable=False)
    item_id = db.Column(db.Integer, nullable=True)  # snapshot reference; menu item may later change
    item_name = db.Column(db.String(120), nullable=False)
    unit_price = db.Column(db.Integer, nullable=False)
    qty = db.Column(db.Integer, nullable=False, default=1)
    sugar = db.Column(db.String(20), nullable=False, default="正常糖")
    ice = db.Column(db.String(20), nullable=False, default="正常冰")
    toppings = db.Column(db.String(300), nullable=False, default="")  # comma-joined names
    note = db.Column(db.String(300), nullable=False, default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def line_total(self):
        return self.unit_price * self.qty

    @property
    def topping_list(self):
        return [t for t in self.toppings.split(",") if t]


class Settlement(db.Model):
    __tablename__ = "settlements"

    id = db.Column(db.Integer, primary_key=True)
    week_id = db.Column(db.Integer, db.ForeignKey("weeks.id"), nullable=False, unique=True)
    actual_total = db.Column(db.Integer, nullable=True)
    note = db.Column(db.Text, nullable=False, default="")
    settled_at = db.Column(db.DateTime, nullable=True)


class MenuTemplate(db.Model):
    __tablename__ = "menu_templates"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    foodpanda_link = db.Column(db.String(500), nullable=False, default="")
    notes = db.Column(db.Text, nullable=False, default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    items = db.relationship(
        "TemplateItem",
        backref="template",
        cascade="all, delete-orphan",
        order_by="TemplateItem.sort_order",
    )
    toppings = db.relationship(
        "TemplateTopping",
        backref="template",
        cascade="all, delete-orphan",
        order_by="TemplateTopping.sort_order",
    )


class TemplateItem(db.Model):
    __tablename__ = "template_items"

    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey("menu_templates.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    price = db.Column(db.Integer, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)


class TemplateTopping(db.Model):
    __tablename__ = "template_toppings"

    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.Integer, db.ForeignKey("menu_templates.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    price = db.Column(db.Integer, nullable=False)
    sort_order = db.Column(db.Integer, nullable=False, default=0)
