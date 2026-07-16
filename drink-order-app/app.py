import json
import os
from datetime import datetime

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for

from auth import requires_admin
from calc import OrderLine, compute_settlement
from models import (
    ICE_OPTIONS,
    SUGAR_OPTIONS,
    Menu,
    MenuItem,
    MenuTemplate,
    MenuTopping,
    Order,
    Settlement,
    TemplateItem,
    TemplateTopping,
    Week,
    db,
)
from seed import seed_default_templates

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def create_app(test_config=None):
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-only-not-secret")
    app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
        "DATABASE_URL", "sqlite:///" + os.path.join(BASE_DIR, "drink_order.db")
    )
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    if test_config:
        app.config.update(test_config)

    db.init_app(app)

    with app.app_context():
        db.create_all()
        seed_default_templates()

    register_routes(app)
    return app


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def get_active_week():
    return Week.query.filter_by(status="active").order_by(Week.id.desc()).first()


def new_week_key():
    now = datetime.now()
    iso_year, iso_week, _ = now.isocalendar()
    base = f"{iso_year}-W{iso_week:02d}"
    candidate = base
    suffix = 1
    while Week.query.filter_by(week_key=candidate).first():
        suffix += 1
        candidate = f"{base}-{suffix}"
    return candidate


def ensure_active_week():
    week = get_active_week()
    if week is None:
        week = Week(week_key=new_week_key(), status="active")
        db.session.add(week)
        db.session.commit()
    return week


def parse_items_json(raw, name_max=120):
    """Parse and validate a JSON list of {name, price} into clean tuples."""
    try:
        data = json.loads(raw or "[]")
    except (TypeError, ValueError):
        raise ValueError("品項資料格式錯誤")
    if not isinstance(data, list):
        raise ValueError("品項資料格式錯誤")
    cleaned = []
    for row in data:
        name = str(row.get("name", "")).strip()[:name_max]
        try:
            price = int(row.get("price"))
        except (TypeError, ValueError):
            raise ValueError(f"品項「{name or '(未命名)'}」價格必須是數字")
        if not name:
            continue
        if price < 0:
            raise ValueError(f"品項「{name}」價格不可為負數")
        cleaned.append((name, price))
    return cleaned


def build_summary_context(week):
    orders = week.orders if week else []
    grand_total_estimate = sum(o.line_total for o in orders)

    item_breakdown = {}
    for o in orders:
        entry = item_breakdown.setdefault(o.item_name, {"qty": 0, "subtotal": 0})
        entry["qty"] += o.qty
        entry["subtotal"] += o.line_total

    person_breakdown = {}
    for o in orders:
        entry = person_breakdown.setdefault(o.person_name, {"orders": [], "subtotal": 0})
        entry["orders"].append(o)
        entry["subtotal"] += o.line_total

    settlement = week.settlement if week else None
    settlement_result = None
    if settlement and settlement.actual_total is not None:
        lines = [
            OrderLine(id=o.id, person_name=o.person_name, unit_price=o.unit_price, qty=o.qty)
            for o in orders
        ]
        settlement_result = compute_settlement(lines, settlement.actual_total)

    return {
        "week": week,
        "orders": orders,
        "grand_total_estimate": grand_total_estimate,
        "item_breakdown": sorted(item_breakdown.items()),
        "person_breakdown": sorted(person_breakdown.items()),
        "settlement": settlement,
        "settlement_result": settlement_result,
    }


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


def register_routes(app):
    # ---------------- Public: order flow ----------------

    @app.route("/")
    def order_page():
        week = get_active_week()
        menu = week.menu if week else None
        published = bool(menu and menu.published)
        closed = False
        if published and menu.deadline and datetime.now() > menu.deadline:
            closed = True
        return render_template(
            "order.html",
            week=week,
            menu=menu,
            published=published,
            closed=closed,
            sugar_options=SUGAR_OPTIONS,
            ice_options=ICE_OPTIONS,
        )

    @app.route("/api/orders", methods=["POST"])
    def create_order():
        week = get_active_week()
        menu = week.menu if week else None
        if not menu or not menu.published:
            return jsonify(ok=False, error="本週尚未發布菜單"), 400
        if menu.deadline and datetime.now() > menu.deadline:
            return jsonify(ok=False, error="已截止收單"), 400

        body = request.get_json(silent=True) or {}
        person_name = str(body.get("person_name", "")).strip()[:80]
        if not person_name:
            return jsonify(ok=False, error="請填寫姓名"), 400

        try:
            item_id = int(body.get("item_id"))
            qty = int(body.get("qty", 1))
        except (TypeError, ValueError):
            return jsonify(ok=False, error="品項或數量格式錯誤"), 400
        if qty < 1:
            return jsonify(ok=False, error="數量至少為 1"), 400

        item = MenuItem.query.filter_by(id=item_id, menu_id=menu.id).first()
        if not item:
            return jsonify(ok=False, error="找不到此品項，請重新選擇"), 400

        sugar = body.get("sugar") or SUGAR_OPTIONS[0]
        if sugar not in SUGAR_OPTIONS:
            sugar = SUGAR_OPTIONS[0]
        ice = body.get("ice") or ICE_OPTIONS[0]
        if ice not in ICE_OPTIONS:
            ice = ICE_OPTIONS[0]

        topping_ids = body.get("topping_ids") or []
        try:
            topping_ids = [int(t) for t in topping_ids]
        except (TypeError, ValueError):
            return jsonify(ok=False, error="加料格式錯誤"), 400
        toppings = []
        if topping_ids:
            toppings = MenuTopping.query.filter(
                MenuTopping.id.in_(topping_ids), MenuTopping.menu_id == menu.id
            ).all()

        note = str(body.get("note", "")).strip()[:300]

        unit_price = item.price + sum(t.price for t in toppings)

        order = Order(
            week_id=week.id,
            person_name=person_name,
            item_id=item.id,
            item_name=item.name,
            unit_price=unit_price,
            qty=qty,
            sugar=sugar,
            ice=ice,
            toppings=",".join(t.name for t in toppings),
            note=note,
        )
        db.session.add(order)
        db.session.commit()

        return jsonify(ok=True, order_id=order.id, line_total=order.line_total)

    @app.route("/summary")
    def summary_page():
        week = get_active_week()
        context = build_summary_context(week)
        context["is_admin"] = False
        return render_template("summary.html", **context)

    # ---------------- Admin: menu setup ----------------

    @app.route("/admin")
    @requires_admin
    def admin_menu():
        week = ensure_active_week()
        menu = week.menu
        templates = MenuTemplate.query.order_by(MenuTemplate.name).all()
        templates_json = [
            {
                "id": t.id,
                "name": t.name,
                "foodpanda_link": t.foodpanda_link,
                "notes": t.notes,
                "items": [{"name": i.name, "price": i.price} for i in t.items],
                "toppings": [{"name": tp.name, "price": tp.price} for tp in t.toppings],
            }
            for t in templates
        ]
        return render_template(
            "admin_menu.html",
            week=week,
            menu=menu,
            templates=templates,
            templates_json=json.dumps(templates_json, ensure_ascii=False),
        )

    @app.route("/admin/publish", methods=["POST"])
    @requires_admin
    def admin_publish():
        week = ensure_active_week()

        store = request.form.get("store", "").strip()[:120]
        foodpanda_link = request.form.get("foodpanda_link", "").strip()[:500]
        notes = request.form.get("notes", "").strip()
        deadline_raw = request.form.get("deadline", "").strip()

        if not store:
            flash("店家名稱為必填", "error")
            return redirect(url_for("admin_menu"))

        deadline = None
        if deadline_raw:
            try:
                deadline = datetime.strptime(deadline_raw, "%Y-%m-%dT%H:%M")
            except ValueError:
                flash("截止時間格式錯誤", "error")
                return redirect(url_for("admin_menu"))

        try:
            items = parse_items_json(request.form.get("items_json"))
            toppings = parse_items_json(request.form.get("toppings_json"))
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("admin_menu"))

        if not items:
            flash("請至少新增一項品項", "error")
            return redirect(url_for("admin_menu"))

        menu = week.menu
        if menu is None:
            menu = Menu(week_id=week.id)
            db.session.add(menu)
        else:
            MenuItem.query.filter_by(menu_id=menu.id).delete()
            MenuTopping.query.filter_by(menu_id=menu.id).delete()

        menu.store = store
        menu.foodpanda_link = foodpanda_link
        menu.notes = notes
        menu.deadline = deadline
        menu.published = True
        db.session.flush()

        for idx, (name, price) in enumerate(items):
            db.session.add(MenuItem(menu_id=menu.id, name=name, price=price, sort_order=idx))
        for idx, (name, price) in enumerate(toppings):
            db.session.add(MenuTopping(menu_id=menu.id, name=name, price=price, sort_order=idx))

        db.session.commit()
        flash("本週菜單已發布", "success")
        return redirect(url_for("admin_menu"))

    @app.route("/admin/reset", methods=["POST"])
    @requires_admin
    def admin_reset():
        old_week = get_active_week()
        if old_week is None:
            flash("目前沒有進行中的一週", "error")
            return redirect(url_for("admin_menu"))

        old_week.status = "archived"

        new_week = Week(week_key=new_week_key(), status="active")
        db.session.add(new_week)
        db.session.flush()

        if old_week.menu:
            old_menu = old_week.menu
            new_menu = Menu(
                week_id=new_week.id,
                store=old_menu.store,
                foodpanda_link=old_menu.foodpanda_link,
                deadline=old_menu.deadline,
                notes=old_menu.notes,
                published=old_menu.published,
            )
            db.session.add(new_menu)
            db.session.flush()
            for item in old_menu.items:
                db.session.add(
                    MenuItem(
                        menu_id=new_menu.id,
                        name=item.name,
                        price=item.price,
                        sort_order=item.sort_order,
                    )
                )
            for topping in old_menu.toppings:
                db.session.add(
                    MenuTopping(
                        menu_id=new_menu.id,
                        name=topping.name,
                        price=topping.price,
                        sort_order=topping.sort_order,
                    )
                )

        db.session.commit()
        flash("已清空本週訂單，開新一輪", "success")
        return redirect(url_for("admin_menu"))

    # ---------------- Admin: menu templates ----------------

    @app.route("/admin/templates", methods=["POST"])
    @requires_admin
    def admin_save_template():
        name = request.form.get("name", "").strip()[:120]
        if not name:
            flash("範本名稱為必填", "error")
            return redirect(url_for("admin_menu"))

        try:
            items = parse_items_json(request.form.get("items_json"))
            toppings = parse_items_json(request.form.get("toppings_json"))
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("admin_menu"))

        template = MenuTemplate(
            name=name,
            foodpanda_link=request.form.get("foodpanda_link", "").strip()[:500],
            notes=request.form.get("notes", "").strip(),
        )
        db.session.add(template)
        db.session.flush()
        for idx, (item_name, price) in enumerate(items):
            db.session.add(
                TemplateItem(template_id=template.id, name=item_name, price=price, sort_order=idx)
            )
        for idx, (topping_name, price) in enumerate(toppings):
            db.session.add(
                TemplateTopping(
                    template_id=template.id, name=topping_name, price=price, sort_order=idx
                )
            )
        db.session.commit()
        flash(f"已另存為範本「{name}」", "success")
        return redirect(url_for("admin_menu"))

    @app.route("/admin/templates/<int:template_id>/delete", methods=["POST"])
    @requires_admin
    def admin_delete_template(template_id):
        template = MenuTemplate.query.get_or_404(template_id)
        db.session.delete(template)
        db.session.commit()
        flash("範本已刪除", "success")
        return redirect(url_for("admin_menu"))

    # ---------------- Admin: summary & settlement ----------------

    @app.route("/admin/summary")
    @requires_admin
    def admin_summary():
        week = get_active_week()
        context = build_summary_context(week)
        context["is_admin"] = True
        return render_template("summary.html", **context)

    @app.route("/admin/settlement", methods=["POST"])
    @requires_admin
    def admin_settlement():
        week = get_active_week()
        if week is None:
            flash("目前沒有進行中的一週", "error")
            return redirect(url_for("admin_summary"))

        raw_total = request.form.get("actual_total", "").strip()
        note = request.form.get("note", "").strip()

        if raw_total == "":
            actual_total = None
        else:
            try:
                actual_total = int(raw_total)
            except ValueError:
                flash("實際結帳金額必須是整數", "error")
                return redirect(url_for("admin_summary"))
            if actual_total < 0:
                flash("實際結帳金額不可為負數", "error")
                return redirect(url_for("admin_summary"))

        settlement = week.settlement
        if settlement is None:
            settlement = Settlement(week_id=week.id)
            db.session.add(settlement)
        settlement.actual_total = actual_total
        settlement.note = note
        settlement.settled_at = datetime.now() if actual_total is not None else None
        db.session.commit()
        flash("已更新結帳金額", "success")
        return redirect(url_for("admin_summary"))

    @app.route("/admin/orders/<int:order_id>/delete", methods=["POST"])
    @requires_admin
    def admin_delete_order(order_id):
        order = Order.query.get_or_404(order_id)
        db.session.delete(order)
        db.session.commit()
        flash("已刪除該筆訂單", "success")
        return redirect(url_for("admin_summary"))

    # ---------------- Admin: history ----------------

    @app.route("/admin/history")
    @requires_admin
    def admin_history():
        weeks = (
            Week.query.filter_by(status="archived").order_by(Week.created_at.desc()).all()
        )
        return render_template("admin_history.html", weeks=weeks)

    @app.route("/admin/history/<int:week_id>")
    @requires_admin
    def admin_history_detail(week_id):
        week = Week.query.get_or_404(week_id)
        context = build_summary_context(week)
        context["is_admin"] = False
        return render_template("admin_history_detail.html", **context)


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
